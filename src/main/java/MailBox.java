import com.sun.mail.imap.IMAPMessage;

import javax.activation.DataSource;
import javax.mail.*;
import javax.mail.internet.MimeBodyPart;
import javax.mail.internet.MimeMessage;
import javax.mail.internet.MimeMultipart;
import javax.mail.search.ComparisonTerm;
import javax.mail.search.ReceivedDateTerm;
import javax.mail.search.SearchTerm;

import org.apache.commons.io.IOUtils;
import org.apache.commons.mail.util.MimeMessageParser;
import org.jsoup.Jsoup;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.*;

public class MailBox {
    public static void main(String[] args) {
        Properties props = System.getProperties();
        props.setProperty("mail.store.protocol", "imaps");
        try {
            Session session = Session.getDefaultInstance(props, null);
            Store store = session.getStore("imaps");
            store.connect("outlook.office365.com", "username", "password");

            Folder inbox = store.getFolder("INBOX");
            inbox.open(Folder.READ_WRITE);

            Calendar calendar = Calendar.getInstance();
            calendar.set( 2022, Calendar.MAY, 31, 2, 0, 0 );
            SearchTerm olderTerm = new ReceivedDateTerm( ComparisonTerm.GE, calendar.getTime() );
            IMAPMessage[] messages = (IMAPMessage[]) inbox.search( olderTerm );
            for ( IMAPMessage message : messages ) {
                // printAllHeader( message );
                downloadAttachment( message );
                System.out.println( getTextFromMessage( message ) );
                // System.out.println( getTextFromMessage( message ) );
                System.out.println("---------------------------------------------------*----------------------------------------");
            }
            store.close();
        }
        catch ( Exception e ) {
            e.printStackTrace();
        }
    }
    public static void downloadAttachment( Message message ) throws MessagingException, IOException {

        Multipart multiPart = (Multipart) message.getContent();
        int numberOfParts = multiPart.getCount();
        for (int partCount = 0; partCount < numberOfParts; partCount++) {
            MimeBodyPart part = (MimeBodyPart) multiPart.getBodyPart(partCount);
            if (Part.ATTACHMENT.equalsIgnoreCase(part.getDisposition())) {
                String file = part.getFileName();
                System.out.println("File name:"+file);
                part.saveFile( "D://test/" + file );

            }
        }
    }
    public static void printAllHeader( Message message ) throws MessagingException {
        Enumeration<Header> headerEnumeration  = message.getAllHeaders();

        while ( headerEnumeration !=null && headerEnumeration.hasMoreElements() ) {
            Header header = headerEnumeration.nextElement();
            System.out.println( header.getName() + " : " + header.getValue() );
        }
    }

    private static String getTextFromMessage( Message message) throws Exception {

        String result = "";
        if (message.isMimeType("text/plain")) {
            result = message.getContent().toString();
        } else if (message.isMimeType("text/html")) {
            result = (String) message.getContent() ;
        } else if (message.isMimeType("multipart/*")) {
            MimeMultipart mimeMultipart = (MimeMultipart) message.getContent();
            result = getTextFromMimeMultipart( mimeMultipart );
        }
        return result;
    }

    public static String getTextFromMimeMultipart( MimeMultipart mimeMultipart ) throws Exception{
        StringBuilder result = new StringBuilder();
        int count = mimeMultipart.getCount();
        for (int i = 0; i < count; i++) {
            BodyPart bodyPart = mimeMultipart.getBodyPart(i);
            if (bodyPart.isMimeType("text/plain")) {
                // result.append("\n").append(bodyPart.getContent());
               // break; // without break same text appears twice in my tests
            } else if (bodyPart.isMimeType("text/html")) {
                String html = (String) bodyPart.getContent();
                result.append(html);
            } else if (bodyPart.getContent() instanceof MimeMultipart){
                result.append( getTextFromMimeMultipart( (MimeMultipart) bodyPart.getContent() ) );
            }
        }
        return result.toString();
    }
}
