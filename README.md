# Exchange邮箱发送邮件

之前都是使用STMP发送邮件，用原来的方法发现不可行

- **利用Exchange Web Services Java API 实现**

-------------------

## Maven

```
<dependency>
    <groupId>com.microsoft.ews-java-api</groupId>
    <artifactId>ews-java-api</artifactId>
    <version>2.0</version>
</dependency>
```

## 代码

```java
import java.net.URI;
import java.net.URISyntaxException;

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import microsoft.exchange.webservices.data.core.ExchangeService;
import microsoft.exchange.webservices.data.core.enumeration.misc.ExchangeVersion;
import microsoft.exchange.webservices.data.core.enumeration.property.BodyType;
import microsoft.exchange.webservices.data.core.service.item.EmailMessage;
import microsoft.exchange.webservices.data.credential.ExchangeCredentials;
import microsoft.exchange.webservices.data.credential.WebCredentials;
import microsoft.exchange.webservices.data.property.complex.MessageBody;

/**
 * Exchange邮件服务工具类
 *
 */
public class ExchangeMailUtil {

  static final Logger LOGGER = LoggerFactory.getLogger(ExchangeMailUtil.class);

  private String mailServer;
  private String user;
  private String password;

  public static void sendMail(String to, String subject, String body) {
    LOGGER.info("send mail to " + to + " with subject " + subject + " and body " + body);
    String mailServer = SettingUtil.getStringProp("MAILSERVER");
    String user = SettingUtil.getStringProp("USER");
    String password = SettingUtil.getStringProp("PASSWORD");
    ExchangeMailUtil mailUtil = new ExchangeMailUtil(mailServer, user, password);
    try {
      mailUtil.send(subject, to, body);
    } catch (Exception e) {
      LOGGER.info("sendMail error {} ", e);
    }
  }

  public ExchangeMailUtil(String mailServer, String user, String password) {
    this.mailServer = mailServer;
    this.user = user;
    this.password = password;
  }

  /**
   * 创建邮件服务
   *
   * @return 邮件服务
   */
  private ExchangeService getExchangeService() {
    ExchangeService service = new ExchangeService(ExchangeVersion.Exchange2007_SP1);
    // 用户认证信息
    ExchangeCredentials credentials;
    credentials = new WebCredentials(user, password);
    service.setCredentials(credentials);
    try {
      service.setUrl(new URI(mailServer));
    } catch (URISyntaxException e) {
      LOGGER.info("getExchangeService error {} ", e);
    }
    return service;
  }

  /**
   * @param subject 邮件标题
   * @param to 收件人列表
   * @param bodyText 邮件内容
   * @throws Exception
   */
  public void send(String subject, String to, String bodyText) throws Exception {
    ExchangeService service = getExchangeService();

    EmailMessage msg = new EmailMessage(service);
    msg.setSubject(subject);
    MessageBody body = MessageBody.getMessageBodyFromText(bodyText);
    body.setBodyType(BodyType.HTML);
    msg.setBody(body);
    msg.getToRecipients().add(to);
    msg.send();
  }

  public static void main(String[] args) throws Exception {
    ExchangeMailUtil mailUtil =
        new ExchangeMailUtil("https://mail.***.com/EWS/exchange.asmx", "用户名", "密码");
    mailUtil.send("Subject", "hsindumas@gmail.com", "content");
    System.out.println("success");
  }
}

```
## 证书
按照上面的方法还是不行，报错：
unable to find valid certification path to requested target 
发现还要安装证书
以下是获取安全证书的一种方法，通过以下程序获取安全证书：

```java
/* 
 * Copyright 2006 Sun Microsystems, Inc.  All Rights Reserved. 
 * 
 * Redistribution and use in source and binary forms, with or without 
 * modification, are permitted provided that the following conditions 
 * are met: 
 * 
 *   - Redistributions of source code must retain the above copyright 
 *     notice, this list of conditions and the following disclaimer. 
 * 
 *   - Redistributions in binary form must reproduce the above copyright 
 *     notice, this list of conditions and the following disclaimer in the 
 *     documentation and/or other materials provided with the distribution. 
 * 
 *   - Neither the name of Sun Microsystems nor the names of its 
 *     contributors may be used to endorse or promote products derived 
 *     from this software without specific prior written permission. 
 * 
 * THIS SOFTWARE IS PROVIDED BY THE COPYRIGHT HOLDERS AND CONTRIBUTORS "AS 
 * IS" AND ANY EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO, 
 * THE IMPLIED WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR 
 * PURPOSE ARE DISCLAIMED.  IN NO EVENT SHALL THE COPYRIGHT OWNER OR 
 * CONTRIBUTORS BE LIABLE FOR ANY DIRECT, INDIRECT, INCIDENTAL, SPECIAL, 
 * EXEMPLARY, OR CONSEQUENTIAL DAMAGES (INCLUDING, BUT NOT LIMITED TO, 
 * PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES; LOSS OF USE, DATA, OR 
 * PROFITS; OR BUSINESS INTERRUPTION) HOWEVER CAUSED AND ON ANY THEORY OF 
 * LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY, OR TORT (INCLUDING 
 * NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE OF THIS 
 * SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE. 
 */  
  
import java.io.BufferedReader;  
import java.io.File;  
import java.io.FileInputStream;  
import java.io.FileOutputStream;  
import java.io.InputStream;  
import java.io.InputStreamReader;  
import java.io.OutputStream;  
import java.security.KeyStore;  
import java.security.MessageDigest;  
import java.security.cert.CertificateException;  
import java.security.cert.X509Certificate;  
  
import javax.net.ssl.SSLContext;  
import javax.net.ssl.SSLException;  
import javax.net.ssl.SSLSocket;  
import javax.net.ssl.SSLSocketFactory;  
import javax.net.ssl.TrustManager;  
import javax.net.ssl.TrustManagerFactory;  
import javax.net.ssl.X509TrustManager;  
  
public class InstallCert {  
  
    public static void main(String[] args) throws Exception {  
        String host;  
        int port;  
        char[] passphrase;  
        if ((args.length == 1) || (args.length == 2)) {  
            String[] c = args[0].split(":");  
            host = c[0];  
            port = (c.length == 1) ? 443 : Integer.parseInt(c[1]);  
            String p = (args.length == 1) ? "changeit" : args[1];  
            passphrase = p.toCharArray();  
        } else {  
            System.out  
                    .println("Usage: java InstallCert <host>[:port] [passphrase]");  
            return;  
        }  
  
        File file = new File("jssecacerts");  
        if (file.isFile() == false) {  
            char SEP = File.separatorChar;  
            File dir = new File(System.getProperty("java.home") + SEP + "lib"  
                    + SEP + "security");  
            file = new File(dir, "jssecacerts");  
            if (file.isFile() == false) {  
                file = new File(dir, "cacerts");  
            }  
        }  
        System.out.println("Loading KeyStore " + file + "...");  
        InputStream in = new FileInputStream(file);  
        KeyStore ks = KeyStore.getInstance(KeyStore.getDefaultType());  
        ks.load(in, passphrase);  
        in.close();  
  
        SSLContext context = SSLContext.getInstance("TLS");  
        TrustManagerFactory tmf = TrustManagerFactory  
                .getInstance(TrustManagerFactory.getDefaultAlgorithm());  
        tmf.init(ks);  
        X509TrustManager defaultTrustManager = (X509TrustManager) tmf  
                .getTrustManagers()[0];  
        SavingTrustManager tm = new SavingTrustManager(defaultTrustManager);  
        context.init(null, new TrustManager[] { tm }, null);  
        SSLSocketFactory factory = context.getSocketFactory();  
  
        System.out  
                .println("Opening connection to " + host + ":" + port + "...");  
        SSLSocket socket = (SSLSocket) factory.createSocket(host, port);  
        socket.setSoTimeout(10000);  
        try {  
            System.out.println("Starting SSL handshake...");  
            socket.startHandshake();  
            socket.close();  
            System.out.println();  
            System.out.println("No errors, certificate is already trusted");  
        } catch (SSLException e) {  
            System.out.println();  
            e.printStackTrace(System.out);  
        }  
  
        X509Certificate[] chain = tm.chain;  
        if (chain == null) {  
            System.out.println("Could not obtain server certificate chain");  
            return;  
        }  
  
        BufferedReader reader = new BufferedReader(new InputStreamReader(  
                System.in));  
  
        System.out.println();  
        System.out.println("Server sent " + chain.length + " certificate(s):");  
        System.out.println();  
        MessageDigest sha1 = MessageDigest.getInstance("SHA1");  
        MessageDigest md5 = MessageDigest.getInstance("MD5");  
        for (int i = 0; i < chain.length; i++) {  
            X509Certificate cert = chain[i];  
            System.out.println(" " + (i + 1) + " Subject "  
                    + cert.getSubjectDN());  
            System.out.println("   Issuer  " + cert.getIssuerDN());  
            sha1.update(cert.getEncoded());  
            System.out.println("   sha1    " + toHexString(sha1.digest()));  
            md5.update(cert.getEncoded());  
            System.out.println("   md5     " + toHexString(md5.digest()));  
            System.out.println();  
        }  
  
        System.out  
                .println("Enter certificate to add to trusted keystore or 'q' to quit: [1]");  
        String line = reader.readLine().trim();  
        int k;  
        try {  
            k = (line.length() == 0) ? 0 : Integer.parseInt(line) - 1;  
        } catch (NumberFormatException e) {  
            System.out.println("KeyStore not changed");  
            return;  
        }  
  
        X509Certificate cert = chain[k];  
        String alias = host + "-" + (k + 1);  
        ks.setCertificateEntry(alias, cert);  
  
        OutputStream out = new FileOutputStream("jssecacerts");  
        ks.store(out, passphrase);  
        out.close();  
  
        System.out.println();  
        System.out.println(cert);  
        System.out.println();  
        System.out  
                .println("Added certificate to keystore 'jssecacerts' using alias '"  
                        + alias + "'");  
    }  
  
    private static final char[] HEXDIGITS = "0123456789abcdef".toCharArray();  
  
    private static String toHexString(byte[] bytes) {  
        StringBuilder sb = new StringBuilder(bytes.length * 3);  
        for (int b : bytes) {  
            b &= 0xff;  
            sb.append(HEXDIGITS[b >> 4]);  
            sb.append(HEXDIGITS[b & 15]);  
            sb.append(' ');  
        }  
        return sb.toString();  
    }  
  
    private static class SavingTrustManager implements X509TrustManager {  
  
        private final X509TrustManager tm;  
        private X509Certificate[] chain;  
  
        SavingTrustManager(X509TrustManager tm) {  
            this.tm = tm;  
        }  
  
        public X509Certificate[] getAcceptedIssuers() {  
            throw new UnsupportedOperationException();  
        }  
  
        public void checkClientTrusted(X509Certificate[] chain, String authType)  
                throws CertificateException {  
            throw new UnsupportedOperationException();  
        }  
  
        public void checkServerTrusted(X509Certificate[] chain, String authType)  
                throws CertificateException {  
            this.chain = chain;  
            tm.checkServerTrusted(chain, authType);  
        }  
    }  
  
}  
```
编译InstallCert.java，然后执行：java InstallCert hostname，比如：
java InstallCert www.twitter.com
会看到如下信息：

```
java InstallCert www.twitter.com  
Loading KeyStore /opt/jdk1.8.0_112/jre/lib/security/cacerts...  
Opening connection to www.twitter.com:443...  
Starting SSL handshake...  
  
javax.net.ssl.SSLHandshakeException: sun.security.validator.ValidatorException: PKIX path building failed: sun.security.provider.certpath.SunCertPathBuilderException: unable to find valid certification path to requested target  
    at com.sun.net.ssl.internal.ssl.Alerts.getSSLException(Alerts.java:150)  
    at com.sun.net.ssl.internal.ssl.SSLSocketImpl.fatal(SSLSocketImpl.java:1476)  
    at com.sun.net.ssl.internal.ssl.Handshaker.fatalSE(Handshaker.java:174)  
    at com.sun.net.ssl.internal.ssl.Handshaker.fatalSE(Handshaker.java:168)  
    at com.sun.net.ssl.internal.ssl.ClientHandshaker.serverCertificate(ClientHandshaker.java:846)  
    at com.sun.net.ssl.internal.ssl.ClientHandshaker.processMessage(ClientHandshaker.java:106)  
    at com.sun.net.ssl.internal.ssl.Handshaker.processLoop(Handshaker.java:495)  
    at com.sun.net.ssl.internal.ssl.Handshaker.process_record(Handshaker.java:433)  
    at com.sun.net.ssl.internal.ssl.SSLSocketImpl.readRecord(SSLSocketImpl.java:815)  
    at com.sun.net.ssl.internal.ssl.SSLSocketImpl.performInitialHandshake(SSLSocketImpl.java:1025)  
    at com.sun.net.ssl.internal.ssl.SSLSocketImpl.startHandshake(SSLSocketImpl.java:1038)  
    at InstallCert.main(InstallCert.java:63)  
Caused by: sun.security.validator.ValidatorException: PKIX path building failed: sun.security.provider.certpath.SunCertPathBuilderException: unable to find valid certification path to requested target  
    at sun.security.validator.PKIXValidator.doBuild(PKIXValidator.java:221)  
    at sun.security.validator.PKIXValidator.engineValidate(PKIXValidator.java:145)  
    at sun.security.validator.Validator.validate(Validator.java:203)  
    at com.sun.net.ssl.internal.ssl.X509TrustManagerImpl.checkServerTrusted(X509TrustManagerImpl.java:172)  
    at InstallCert$SavingTrustManager.checkServerTrusted(InstallCert.java:158)  
    at com.sun.net.ssl.internal.ssl.JsseX509TrustManager.checkServerTrusted(SSLContextImpl.java:320)  
    at com.sun.net.ssl.internal.ssl.ClientHandshaker.serverCertificate(ClientHandshaker.java:839)  
    ... 7 more  
Caused by: sun.security.provider.certpath.SunCertPathBuilderException: unable to find valid certification path to requested target  
    at sun.security.provider.certpath.SunCertPathBuilder.engineBuild(SunCertPathBuilder.java:236)  
    at java.security.cert.CertPathBuilder.build(CertPathBuilder.java:194)  
    at sun.security.validator.PKIXValidator.doBuild(PKIXValidator.java:216)  
    ... 13 more  
  
Server sent 2 certificate(s):  
  
 1 Subject CN=www.twitter.com, O=example.com, C=US  
   Issuer  CN=Certificate Shack, O=example.com, C=US  
   sha1    2e 7f 76 9b 52 91 09 2e 5d 8f 6b 61 39 2d 5e 06 e4 d8 e9 c7   
   md5     dd d1 a8 03 d7 6c 4b 11 a7 3d 74 28 89 d0 67 54   
  
 2 Subject CN=Certificate Shack, O=example.com, C=US  
   Issuer  CN=Certificate Shack, O=example.com, C=US  
   sha1    fb 58 a7 03 c4 4e 3b 0e e3 2c 40 2f 87 64 13 4d df e1 a1 a6   
   md5     72 a0 95 43 7e 41 88 18 ae 2f 6d 98 01 2c 89 68   
  
Enter certificate to add to trusted keystore or 'q' to quit: [1]  
```
输入1，回车，然后会在当前的目录下产生一个名为“ssecacerts”的证书。
将证书拷贝到$JAVA_HOME/jre/lib/security目录下，或者通过以下方式：
System.setProperty("javax.Net.ssl.trustStore", "你的jssecacerts证书路径");


注意：因为是静态加载，所以要重新启动你的Web Server，证书才能生效。