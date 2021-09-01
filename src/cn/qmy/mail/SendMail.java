package cn.qmy.mail;

import java.util.Properties;

import javax.mail.Message;
import javax.mail.Message.RecipientType;
import javax.mail.Session;
import javax.mail.Transport;
import javax.mail.internet.InternetAddress;
import javax.mail.internet.MimeMessage;

public class SendMail {
	public static void main(String[] args) {
		SendMail s = new SendMail();
		// 参数为：收件人的邮箱地址
		s.sendMails("dyxiye@163.com","32sdsg");
	}

	public boolean sendMails(String recipients,String code) {
		Email mail = new Email();
		// 发件人的邮箱地址（要完整），会显示在收件人的邮件中
		mail.setSender("dyxiye@163.com");
		// 发件人登录邮箱的账号（@符合前面的部分）
		mail.setUserName("dyxiye");
		// 下面填的是邮箱客户端授权码，切忌：邮箱务必要开启（POP3/SMTP服务）
		mail.setPassword("dingye1995");
		try {
			// 创建邮件对象
			Session session = null;
			Properties props = new Properties();
			// 此处为发送方邮件服务器地址，要根据邮箱的不同需要自行设置
			props.put("mail.smtp.host", "smtp.163.com");
			props.setProperty("mail.transport.protocol", "smtp");
			// SMTP端口号
			props.put("mail.smtp.port", "25");
			// 设置成需要邮件服务器认证
			props.put("mail.smtp.auth", "true");
			props.put("mail.debug", "true");
			session = Session.getInstance(props);
			session.setDebug(true);
			Message message = new MimeMessage(session);
			// 设置发件人
			message.setFrom(new InternetAddress(mail.getSender()));
			// 设置收件人
			message.addRecipient(RecipientType.TO, new InternetAddress(recipients));
			// 设置标题
			message.setSubject("甜儿验证码");
			// 邮件内容，根据自己需要自行制作模板
			message.setContent("<p>尊敬的xiye用户：</p><p>您好!</p><p>您的验证码是："+code+"。</p>"
					, "text/html;charset=utf-8");
			// 发送邮件
			Transport transport = session.getTransport("smtp");
			transport.connect("smtp.163.com", mail.getUserName(), mail.getPassword());// 以smtp方式登录邮箱
			// 发送邮件,其中第二个参数是所有已设好的收件人地址
			transport.sendMessage(message, message.getAllRecipients());
			transport.close();
			return true;
		} catch (Exception e) {
			e.printStackTrace();
			return false;
		}
	}
}
