DECLARE
  k_host            CONSTANT VARCHAR2(100) := 'podNNNNN.outlook.com';
  k_port            CONSTANT INTEGER       := 587;
  k_wallet_path     CONSTANT VARCHAR2(100) := 'file:/home/oracle/wallets';
  k_wallet_password CONSTANT VARCHAR2(100) := 'XXXXXXXX';
  k_domain          CONSTANT VARCHAR2(100) := 'localhost';
  k_username        CONSTANT VARCHAR2(100) := 'XXXXXXXX';
  k_password        CONSTANT VARCHAR2(100) := 'XXXXXXXX';
  k_sender          CONSTANT VARCHAR2(100) := 'XXXXXXXX';
  k_recipient       CONSTANT VARCHAR2(100) := 'XXXXXXXX';
  k_subject         CONSTANT VARCHAR2(100) := 'Test TLS mail';
  k_body            CONSTANT VARCHAR2(100) := 'Message body';
  
  l_conn    utl_smtp.connection;
  l_reply   utl_smtp.reply;
  l_replies utl_smtp.replies;
BEGIN
  dbms_output.put_line('utl_smtp.open_connection');
  
  l_reply := utl_smtp.open_connection
             ( host                          => k_host
             , port                          => k_port
             , c                             => l_conn
             , wallet_path                   => k_wallet_path
             , wallet_password               => k_wallet_password
             , secure_connection_before_smtp => FALSE
             );
 
  IF l_reply.code != 220
  THEN
    raise_application_error(-20000, 'utl_smtp.open_connection: '||l_reply.code||' - '||l_reply.text);
  END IF;
 
  dbms_output.put_line('utl_smtp.ehlo');
  
  l_replies := utl_smtp.ehlo(l_conn, k_domain);
  
  FOR ri IN 1..l_replies.COUNT
  LOOP
    dbms_output.put_line(l_replies(ri).code||' - '||l_replies(ri).text);
  END LOOP;
 
  dbms_output.put_line('utl_smtp.starttls');
  
  l_reply := utl_smtp.starttls(l_conn);
 
  IF l_reply.code != 220
  THEN
    raise_application_error(-20000, 'utl_smtp.starttls: '||l_reply.code||' - '||l_reply.text);
  END IF;
 
  dbms_output.put_line('utl_smtp.ehlo');
  
  l_replies := utl_smtp.ehlo(l_conn, k_domain);
  
  FOR ri IN 1..l_replies.COUNT
  LOOP
    dbms_output.put_line(l_replies(ri).code||' - '||l_replies(ri).text);
  END LOOP;
  
  dbms_output.put_line('utl_smtp.auth');
  
  l_reply := utl_smtp.auth(l_conn, k_username, k_password, utl_smtp.all_schemes);
 
  IF l_reply.code != 235
  THEN
    raise_application_error(-20000, 'utl_smtp.auth: '||l_reply.code||' - '||l_reply.text);
  END IF;

  dbms_output.put_line('utl_smtp.mail');
  
  l_reply := utl_smtp.mail(l_conn, k_sender);
  
  IF l_reply.code != 250
  THEN
    raise_application_error(-20000, 'utl_smtp.mail: '||l_reply.code||' - '||l_reply.text);
  END IF;
 
  dbms_output.put_line('utl_smtp.rcpt');
  
  l_reply := utl_smtp.rcpt(l_conn, k_recipient);
  
  IF l_reply.code NOT IN (250, 251)
  THEN
    raise_application_error(-20000, 'utl_smtp.rcpt: '||l_reply.code||' - '||l_reply.text);
  END IF;
 
  dbms_output.put_line('utl_smtp.open_data');
  
  l_reply := utl_smtp.open_data(l_conn);
  
  IF l_reply.code != 354
  THEN
    raise_application_error(-20000, 'utl_smtp.open_data: '||l_reply.code||' - '||l_reply.text);
  END IF;
  
  dbms_output.put_line('utl_smtp.write_data');
  
  utl_smtp.write_data(l_conn, 'From: '||k_sender||utl_tcp.crlf);
  utl_smtp.write_data(l_conn, 'To: '||k_recipient||utl_tcp.crlf);
  utl_smtp.write_data(l_conn, 'Subject: '||k_subject||utl_tcp.crlf);
  utl_smtp.write_data(l_conn, utl_tcp.crlf||k_body);
 
  dbms_output.put_line('utl_smtp.close_data');
  
  l_reply := utl_smtp.close_data(l_conn);
  
  IF l_reply.code != 250
  THEN
    raise_application_error(-20000, 'utl_smtp.close_data: '||l_reply.code||' - '||l_reply.text);
  END IF;
  
  dbms_output.put_line('utl_smtp.quit');
  
  l_reply := utl_smtp.quit(l_conn);
  
  IF l_reply.code != 221
  THEN
    raise_application_error(-20000, 'utl_smtp.quit: '||l_reply.code||' - '||l_reply.text);
  END IF;

EXCEPTION
  WHEN    utl_smtp.transient_error
       OR utl_smtp.permanent_error
  THEN
    BEGIN
      utl_smtp.quit(l_conn);
    EXCEPTION
      WHEN    utl_smtp.transient_error
           OR utl_smtp.permanent_error
      THEN
        NULL;
    END;
    
    raise_application_error(-20000, 'Failed to send mail due to the following error: '||SQLERRM);
    
END;
/
