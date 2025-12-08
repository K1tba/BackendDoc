#!/usr/bin/perl

use strict;
use warnings;
use utf8;

use File::Path qw();
use File::Spec;

use File::Basename;
use FindBin qw($RealBin);
use lib dirname($RealBin);
use lib dirname($RealBin) . '/Kernel/cpan-lib';
use lib dirname($RealBin) . '/Custom';

use Kernel::System::ObjectManager;

local $Kernel::OM = Kernel::System::ObjectManager->new(
    'Kernel::System::Log' => {
        LogPrefix => 'OTRS-otrs.ExportEmails.pl',
    },
);

# --
# CONFIGURATION!!!
# --
my $ExportBaseDir   = '/opt/otrs/var/article/otrs_export';    # Целевая сетевая папка
my $TargetQueueName = 'Raw';                                  # Имя очереди (точно как в OTRS)
my $StartDateTime   = '2025-01-01 00:00:00';                  # Начало периода (включительно)
my $EndDateTime     = '2025-12-31 23:59:59';                  # Конец периода (включительно)
# --

# init OTRS object manager
local $Kernel::OM;
$Kernel::OM = Kernel::System::ObjectManager->new();

# get needed objects
my $ArticleObject = $Kernel::OM->Get('Kernel::System::Ticket::Article');
my $LayoutObject  = $Kernel::OM->Get('Kernel::Output::HTML::Layout');
my $MainObject    = $Kernel::OM->Get('Kernel::System::Main');
my $QueueObject   = $Kernel::OM->Get('Kernel::System::Queue');
my $TicketObject  = $Kernel::OM->Get('Kernel::System::Ticket');


# get queue name
my $QueueID = $QueueObject->QueueLookup( Queue => $TargetQueueName );
if ( !$QueueID ) {
    die "Ошибка: очередь '$TargetQueueName' не найдена в OTRS.\n";
}

print "Найдена очередь: $TargetQueueName (ID: $QueueID)\n";

# ticket search
my @TicketIDs = $TicketObject->TicketSearch(
    Result                    => 'ARRAY',
    QueueIDs                  => [ $QueueID ],
    TicketCreateTimeNewerDate => $StartDateTime,
    TicketCreateTimeOlderDate => $EndDateTime,
    UserID                    => 1,
);

print "Найдено заявок: " . scalar(@TicketIDs) . "\n";

for my $TicketID ( @TicketIDs ) {

    # get ticket data
    my %Ticket = $TicketObject->TicketGet(
        TicketID => $TicketID,
        UserID   => 1,
    );

    my $QueueName    = $Ticket{Queue};
    my $TicketNumber = $Ticket{TicketNumber};

    # get all articles ticket (filter at communication channel - only  email)
    my @Articles = $ArticleObject->ArticleList(
        TicketID             => $TicketID,
        CommunicationChannel => 'Email', 
    );

    ARTICLE:
    for my $MetaArticle ( @Articles ) {

        my $ArticleID     = $MetaArticle->{ArticleID};
        my $ArticleNumber = $MetaArticle->{ArticleNumber};

        my $ArticleBackendObject = $ArticleObject->BackendForArticle(
            TicketID  => $TicketID,
            ArticleID => $ArticleID,
        );

        # skip if sender type - system
        next ARTICLE if $ArticleBackendObject->ChannelNameGet() ne 'Email';

        # get article data
        my %Article = $ArticleBackendObject->ArticleGet(
            ArticleID => $ArticleID,
            UserID    => 1,
        );

        # extract key fields
        my $From       = $Article{From}    || 'unknown';
        my $Subject    = $Article{Subject} || '(no subject)';
        my $SenderType = $Article{SenderType};
        my $CreateTime = $Article{CreateTime};

        # subject length limit
        $Subject = substr($Subject, 0, 80);

        # build filename
        my $Filename = sprintf(
            "%d-%s(%s)-%s-%s.eml",
            $ArticleNumber,
            $From,
            $SenderType,
            $Subject,
            $CreateTime
        );

        # perform FilenameCleanup here already to check for
        #   conflicting existing attachment files correctly
        $Filename = $MainObject->FilenameCleanUp(
            Filename => $Filename,
            Type     => 'Local',
        );

        # path build : ExportBaseDir/QueueName/TicketNumber/
        my $TargetDir = File::Spec->catdir($ExportBaseDir, $QueueName, $TicketNumber);
        make_path($TargetDir, { error => \my $Error });
        if (@$Error) {
            warn "Не удалось создать каталог $TargetDir: $Error->[0]->{message}\n";
            next;
        }

        my $FilePath = File::Spec->catfile($TargetDir, $Filename);

        # get article plain text
        my $Plain = $ArticleBackendObject->ArticlePlain(
            TicketID  => $TicketID,
            ArticleID => $ArticleID
        );

        # get raw-content message in format .eml
        my $EMLContent = $LayoutObject->Attachment(
            Filename    => $Filename,
            ContentType => 'message/rfc822',
            Content     => $Plain,
            Type        => 'attachment',
        );

        unless ( defined $EMLContent && length $EMLContent ) {
            warn "Пустое или отсутствующее содержимое для статьи $ArticleID (заявка $TicketID)\n";
            next;
        }

        # write file
        open my $FH, '>:raw', $FilePath or do {
            warn "Ошибка записи в $FilePath: $!\n";
            next;
        };
        print $FH $EMLContent;
        close $FH;

        print "Сохранено: $FilePath\n";
    }
}

print "Экспорт завершён.\n";
