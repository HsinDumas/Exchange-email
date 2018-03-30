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
   * @param subject  邮件标题
   * @param to       收件人列表
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
    ExchangeMailUtil mailUtil = new ExchangeMailUtil("https://mail.***.com/EWS/exchange.asmx", "用户名", "密码");
    mailUtil.send("Subject", "hsindumas@gmail.com", "content");
    System.out.println("success");
  }
}