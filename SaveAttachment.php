<?
class SaveAttachment {

    private $mail_login;
    private $mail_password;
    private $host_imap;
    private $connection;
    

    public function __construct($email, $password, $host = 'yandex.ru') {   // gmail.com
        $this->mail_login = $email;
        $this->mail_password = $password;
        $this->host_imap = "{imap.$host:993/imap/ssl}";
        $this->connection = imap_open($this->host_imap, $this->mail_login, $this->mail_password);
        if(!$this->connection)
            exit('Ошибка соединения');
    }

    private function getAttachment($str) {

        $msg_num = imap_search($this->connection, $str);  // письма
        if($msg_num && $msg_num != []) {
            $email_attachments = [];
            foreach ($msg_num as $emailsNumber) {
                $msg_header = imap_header($this->connection, $emailsNumber);
    //            echo $msg_header->MailDate;  // дата
    //            echo(imap_mime_header_decode($msg_header->subject)[0]->text);  // тема письма
                $attachments = [];
                $msg_structure = imap_fetchstructure($this->connection, $emailsNumber);

                if (isset($msg_structure->parts)) {

                    for ($i = 0; $i < count($msg_structure->parts); $i++) {

                        $attachments[$i] = array(
                            'is_attachment' => false,
                            'filename' => '',
                            'name' => '',
                            'attachment' => '',
                            'date' => $msg_header->MailDate
                        );

                        if ($msg_structure->parts[$i]->ifdparameters) {
                            foreach ($msg_structure->parts[$i]->dparameters as $object) {
                                if (strtolower($object->attribute) == 'filename') {
                                    $attachments[$i]['is_attachment'] = true;
                                    $attachments[$i]['filename'] = imap_mime_header_decode($object->value)[0]->text;
                                }
                                if (strtolower($object->attribute) == 'name') {
                                    $attachments[$i]['is_attachment'] = true;
                                    $attachments[$i]['name'] = imap_mime_header_decode($object->value)[0]->text;
                                }
                            }
                        }

                        if ($attachments[$i]['is_attachment']) {
                            $attachments[$i]['attachment'] = imap_fetchbody($this->connection, $emailsNumber, $i + 1);
                            if ($msg_structure->parts[$i]->encoding == 3) {
                                $attachments[$i]['attachment'] = base64_decode($attachments[$i]['attachment']);
                            } elseif ($msg_structure->parts[$i]->encoding == 4) {
                                $attachments[$i]['attachment'] = quoted_printable_decode($attachments[$i]['attachment']);
                            }
                        }
                    }
                    $email_attachments = array_merge($email_attachments, array_filter($attachments, function ($attachment) {
                        return $attachment['is_attachment'] == 1;
                    }));
                }
            }
            if ($email_attachments && $email_attachments != [])
                return $email_attachments;
            else
                echo 'Вложений нет';
        }else {
            echo 'Писем нет';
        }
        return false;
    }
    private function writeFile($name, $data) {
        $path = $_SERVER['DOCUMENT_ROOT'].'/imap/xlsx/';
        $this->clear_dir($path);

        $path .= $name;
        $fp = fopen($path, "w");
        fwrite($fp, $data);
        fclose($fp);
        return $path;
    }

    private function clear_dir($dir) {
        $list = scandir($dir);
        unset($list[0],$list[1]);
        foreach ($list as $file)
        {
            unlink($dir.$file);
        }
    }

    public function createFile( $str = 'ALL') {  // 'UNSEEN' - непрочитанные
        $arrPath = [];
        $result = $this->getAttachment($str);
        foreach ($result as $attachment) {
//            echo $attachment['filename'];
            $path = $this->writeFile($attachment['filename'], $attachment['attachment']);
            $arrPath[$attachment['date']] = $path;

            $fp = fopen($_SERVER['DOCUMENT_ROOT'] . '/imap/date_update.txt', "w");
            fwrite($fp, $attachment['date']);
            fclose($fp);
        }
        return $arrPath;
    }
}

