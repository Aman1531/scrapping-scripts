#!/usr/bin/perl -w 
print "Content-type: text/plain\n\n";
use strict;
use WWW::Mechanize;
use Data::Dumper;
use HTTP::Cookies;
use  Spreadsheet::ParseExcel;
use DBI;

my $cookie_jar=HTTP::Cookies->new(
    'file' => 'cookies2.lwp',  # where to read/write cookies
   'autosave' => 1,  
   );

my $dsn = "DBI:mysql:greensee_booking;host=localhost";
my $dbh = DBI->connect($dsn, 'greensee_booker', '@1531@xxx');

    my $login    = "1823932";
    my $password = "mariposa27";


my $url = "https://admin.booking.com";
my $user_agent ="Mozilla/5.0 (X11; Linux armv7l) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/56.0.2924.84 Safari/537.36";


my $mech = WWW::Mechanize->new(agent=>$user_agent);
$mech->cookie_jar($cookie_jar);
$mech->get($url);
my $form = $mech->form_number(2);
use HTML::Form;
my $input = HTML::Form::Input->new(name=>'csrf_token');
$input->add_to_form($form);
      $mech->submit_form(
          fields      => { loginname => $login ,password => $password,csrf_token=>'empty-token' },
      );

my  @links=$mech->links;
 my  $logout ='';
 my  $res ='';
 for my $link (@links) {
     $_ = $link->url;
   $logout=$link if m/logout/;
   $res=$link if m/reservation/;

    
}


$mech->get($res);
die unless ($mech->success);


my $dlink =$mech->find_link( url_regex => qr/download/i );

$mech->mirror($dlink->url_abs,'latest.xls') if $dlink;
 
my $parser=Spreadsheet::ParseExcel->new();
my $workbook=$parser->parse('latest.xls');

if(!defined $workbook)
{
die $parser->error(),".\n";
}

my $worksheet=$workbook->worksheet('Sheet1');

my @data;

    my ( $row_min, $row_max ) = $worksheet->row_range();
    my ( $col_min, $col_max ) = $worksheet->col_range();

    for my $row ( $row_min .. $row_max ) {
        my @values;
        for my $col ( $col_min .. $col_max ) {
            my $cell = $worksheet->get_cell( $row, $col );
            next unless $cell;
	    push @values , $cell->value();
      	    push @values , undef if $col eq 0 or $col eq 7;
        }
     
	push @data, \@values;

    }

my $hdr = shift @data;

local $SIG{__WARN__} = sub {
};

my @records ;
for my $row (@data) { 
    my   $sth=  $dbh->prepare("INSERT INTO reservation  VALUES (?, ?,?,?,?,?,?, ?,?,?,?,?,?,?,?)");
       
    $sth->execute(@$row)  && push(@records, @$row[0]);

}

use WWW::Mechanize::Link;

if(@records){

     my $glink =  'https://admin.booking.com/hotel/hoteladmin/extranet_ng/manage/booking.html?hotel_id=1823932&lang=en';

     $_=$mech->uri;
     my $ses_id =  $1 if /ses=([^&]+)/;

     $glink .= '&ses='.$ses_id.'&res_id=123';

    my @emails ;
     for my $num (@records) {
    my @rec;

        $glink =~ s/res_id=\d+/res_id=$num/;

my    $u = WWW::Mechanize::Link->new( {  url  => $glink});

    $mech->get($u);

    my $email = $mech->find_link(tag =>'a' , id=>'email');
    push @rec, $num;
    
    $email? push @rec,$email->text:push @rec, undef;
     
         $_=$mech->text;
    push @rec,   $1 if /Room:.+?\((.+?)\)/;
    push @emails , \@rec;


  
     }
     
#     print Dumper @emails;
for my $row (@emails) { 
    my   $sth=  $dbh->prepare("UPDATE reservation SET email=? , room_type=? WHERE book_number=?");
 $sth->execute( @$row[1],@$row[2], @$row[0]); 
}


}

$mech->get($logout);
die unless ($mech->success);
