#!/usr/bin/perl

# Reads mail from Inbox, moves parsed messges to a Parsed folder, prints a summary csv.
# Only reads messages whose subject contains the plant specified by the --plant arg.
# 
# https://ask.metafilter.com/353374/Parsing-plant-orders-with-Perl
#
# Tested on Linux
# Change email host/user on lines 50ish

use strict;
use warnings;
use Mail::IMAPClient; # apt-get install libmail-imapclient-perl
use Data::Dumper;
use Getopt::Long;

main();
exit;

sub main {
    my ($debug, $only_unread, $password, $plant, $save, $since, $test);

    GetOptions(
        "debug" => \$debug,
        "only-unread" => \$only_unread,
        "password=s" => \$password,
        "plant=s" => \$plant,
        "save=s" => \$save,
        "since=s" => \$since,
        "test" => \$test,
    );

    if ( not $plant ) {
        usage("Missing plant");
        return;
    }
    if ( $test ) {
        test_parsing($debug, $plant);
        return;
    }
    if ( not $password ) {
        usage("Missing password");
        return;
    } elsif ( $since  && $since !~ /^\d{1,2}-(?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)-\d{4}$/ ) {
        print "Expected date format like 31-Jan-2021\n";
        return;
    } elsif ( @ARGV ) {
        usage("Unrecognized args (".join(' ', @ARGV).')');
        return;
    }
    my $move_to = 'Parsed';

    my $imap = init_imap(
        host => 'mail.example.com',
        port => 143,
        user => 'user@example.com',
        pass => $password,
        folder => 'Inbox',
    );

    my @search_args;

    if ( $only_unread ) {
        push @search_args, 'UNSEEN';
    }
    if ( $since ) {
        push @search_args, ( SINCE => $since );
    }
    push @search_args, ( 'SUBJECT', $plant );

    process( $imap, $move_to, $plant, $save, \@search_args );
    return;
}

sub usage {
    my $msg = shift;
    if ( $msg ) {
        print "$msg\n";
    }
    print "usage: $0 --password=secret --plant=daisy [--save=output.csv] [--since=31-Jan-2021] [--only-unread]\n";
    print "usage: $0 --test --plant='Olearia traversii' [--debug]\n";
    return;
}

sub init_imap {
    my (%args) = @_;

    my $imap = Mail::IMAPClient->new(
        Debug => 0,
        Server => $args{host},
        Port => $args{port}||143,
        User => $args{user},
        Password => $args{pass},
        Timeout => $args{timeout}||120,
    ) || die "Couldn't create IMAP client";

    $imap->select($args{folder}) ||
        die "Can't select IMAP folder" . $args{folder} . ':' . $!;
    return $imap;
}

sub search {
    my ($imap, $search_args) = @_;
    my @search_args = @{$search_args} ? @{$search_args} : ('ALL',);
    # see 6.4.4. in http://www.faqs.org/rfcs/rfc2060.html 
    # e.g. ('ALL'), ('FROM', '"someone@there"'), ('SUBJECT', '"yadda"'), ('BODY', '"needle"')
    return [$imap->search( @search_args )] || [];
}

sub get_msg {
    my ($imap, $msg_id) = @_;

    my $msg_data = $imap->parse_headers( $msg_id, "Date", "From", "To", "Subject" );

    for my $key (keys %{$msg_data}) {
        if ( 'ARRAY' eq ref $msg_data->{$key} ) {
            $msg_data->{$key} = join("\n", @{$msg_data->{$key}} );
        }
    }

    my $body = $imap->body_string( $msg_id );
    $body =~ s/\r\n/\n/smg;
    $msg_data->{body} = $body;

    $msg_data->{msg_id} = $msg_id;

    $msg_data->{printable_msg} =
        "From: "         . $msg_data->{From}            . "\n" .
        "To: "           . $msg_data->{To}              . "\n" .
        "Date: "         . $msg_data->{Date}            . "\n" .
        "Subject: "      . $msg_data->{Subject}         . "\n" .
        "\n" .
        $msg_data->{body} . "\n";

    return $msg_data;
}

sub process {
    my ( $imap, $move_to, $plant, $save, $search_args ) = @_;
    my $summary = {};

    my $msg_ids = search( $imap, $search_args );
    print "Found " . @{$msg_ids} . ' message' . ( @{$msg_ids} == 1 ? '' : 's' ) . "\n";

    # ticker to display dots so you know it's not stuck.
    my ($ticker_n, $ticker_i, $need_newline) = (10,10,0);
    for my $msg_id ( @{$msg_ids} ) {

        my $flags = $imap->flags( $msg_id );
        next if grep /\\Deleted/, @{$flags};

        my $msg_data = get_msg( $imap, $msg_id );

        if ( should_move_msg( $msg_data, $plant, $summary ) ) {
            $imap->move( $move_to, [$msg_id] );
            $imap->expunge();
        } else {
            # if it wasn't seen before, then restore its 'unseen' status.
            if ( not grep /\\Seen/, @{$flags} ) {
                $imap->unset_flag( '\Seen', $msg_id );
            }
        }

        if ( --$ticker_i <= 0 ) {
            print ".";
            $ticker_i = $ticker_n;
            $need_newline = 1;
        }
    }

    print "\n" if ( $need_newline );
    summarize( $plant, $save, $summary );
    return;
}

sub summarize {
    my ( $plant, $save, $summary ) = @_;
    return if not %{$summary};

    my $fh;
    if ( $save ) {
        open($fh, '>', $save) || die "Could not create $save for writing. $!";
    } else {
        $fh = \*STDOUT;
    }
    my $csv_header = csvify( [ qw(have count price plant date from) ] ) . "\n";
    print {$fh} $csv_header;

    # iterate over messages by sender and then msg_id
    for my $key (sort
        { 
            ($summary->{$a}->{msg_data}->{From}//'') cmp ($summary->{$b}->{msg_data}->{From}//'')
            ||
            $a <=> $b
        } keys %{$summary}
    ) {
        my $row = $summary->{$key};
        print {$fh} csvify( [
            $row->{have}//'',
            $row->{count}//'',
            $row->{price}//'',
            $plant,
            $row->{msg_data}->{Date}//'',
            $row->{msg_data}->{From}//'',
        ] ) . "\n";
    }
    if ( $save ) {
        close($fh) || die "Faield to close $save. $!";
    }
    return;
}

sub csvify {
    my $row = shift;
    return '"' .  join('","', @{$row}) . '"';
}

sub should_move_msg {
    my ($msg_data, $plant, $summary) = @_;

    my $move_it = 0;

    # if plant is not in subject, return, do not move mail.
    if ( $msg_data->{Subject} !~ /$plant/ ) {
        return $move_it;
    }

    my $body = $msg_data->{body};
    my $msg_id = $msg_data->{msg_id};

    my ($have, $count, $price);

    if ($body =~ /have (\d+)/i ) {
        $have = 'Y';
        $count = $1;
    } elsif ( $body =~ /(\d+)\s+(?:x\s+)?$plant/i ) { # accept "123 x $plant" and "123 $plant"
        $have = 'Y';
        $count = $1;
    } elsif ( $body =~ /$plant\s+(?:x\s*)?(\d+)/i ) { # accept "$plant x 123" and "$plant 123"
        $have = 'Y';
        $count = $1;
    }
    if ( $body =~ /We can/i ) {
        $have = 'Y';
    }
    if ($body =~ /(?:don't|do not|cannot|can't)/i ) {
        $have = 'N';
    } elsif ($body =~ /(?:don’t|can’t)/i ) {
        $have = 'N';
    }
    if ( defined $have ) {

        if ( $have eq 'Y' ) {
            if ($body =~ /\$\s*(\d+(?:\.\d*)?|\.(?:\d+))/ ) {
                $price = $1;
            }
        }
        $move_it = 1;
        $summary->{$msg_id} = {
            msg_data => $msg_data,
            have => $have,
            count => ($count//''),
            price => ($price//''),
        };
    }
    return $move_it;
}

sub test_parsing {
    my ($debug, $plant) = @_;

    STDOUT->autoflush(1);
    STDERR->autoflush(1);

    while(<DATA>) {
        chomp;
        my $line = $_;
        my ($expect_move, $expect_have, $expect_count, $expect_price, @rest) = split(',', $line);
        my $body = join(',', @rest);
        my $msg_id = 123;
        my $msg_data = {
            Subject => "Re: $plant availability",
            body => $body,
            msg_id => $msg_id,
        };
        my $summary = {};
        my $move_it = should_move_msg($msg_data, $plant, $summary);
        my $move_it_yn = $move_it ? 'Y' : 'N';
        my $details = $summary->{$msg_id};
        
        if ( $debug ) {
            print "\n----\n";
            print 'line          is ('.$line.")\n";
            print 'details       is ('.Dumper($details).")\n";
            print 'expect_have   is ('.$expect_have.")\n";
            print 'expect_count  is ('.$expect_count.")\n";
            print 'expect_price  is ('.$expect_price.")\n";
        }

        if ( $expect_move ne $move_it_yn ) {
            print "FAIL: Mismatch move_it=$move_it_yn expected=$expect_move for $body\n";
            next;
        }
        if ( $expect_move eq 'N' && not defined $details ) {
            next;
        }
        if ( $expect_have ne ($details->{have}//'x') ) {
            print "FAIL: Mismatch have=$details->{have} expected=$expect_have for $body\n";
        } elsif ( $expect_count ne ($details->{count}//'x') ) {
            print "FAIL: Mismatch count=$details->{count} expected=$expect_count for $body\n";
        } elsif ( $expect_price ne ($details->{price}//'x') ) {
            print "FAIL: Mismatch price=$details->{price} expected=$expect_price for $body\n";
        } else {
            print "PASS: $line\n";
        }
    }
    return;
}

__END__
Y,Y,21,,The PB 6.5’s appear to have 21 but I may have missed one tucked somewhere. Pic of both attached
Y,Y,,6.35,We can do this many plants for you but only in a 2.5 litre grade @ $6.35 + gst each.
Y,Y,45,,We can indeed supply you with 45 x Olearia Traversii.
Y,Y,45,9,Hi Nigel. We have 45 Olearia traversii. They are about a metre tall in a PB5 bag. Price would be $9 plus gst.
Y,Y,45,3.50,Yes we have 45 Olearia traversii but unfortunately only V150 grade ($3.50 + GST each), these are good plants though and are probably 75 cm tall.
Y,N,,,Sorry I don’t have any Traversii available I do have O dartonii Pb3 $5.89 plus GST and freight
N,,,,Some unparsable message
