using System;
using System.Net;
using System.Net.Mail;
using System.Text;

namespace TshRet
{
	public class CNotification
	{
		private string	m_sSmtp;
		private int		m_iPort;
		private string	m_sUserName;
		private string	m_sUserId;
		private string	m_sPassword;
		private bool	m_bSslTls;

		public CNotification(CSettings.CContents scontents)
		{
			m_sSmtp			= scontents.MailSmtp;
			m_iPort			= scontents.MailSmtpPort;
			m_sUserName		= scontents.MailUserName;
			m_sUserId		= scontents.MailUserId;
			m_sPassword		= scontents.MailPassword;
			m_bSslTls		= scontents.MailSslTls;
		}

		~CNotification()
		{
			m_sSmtp			= null;
			m_sUserId		= null;
			m_sUserName		= null;
			m_sPassword		= null;
		}

		public void SendNotificationMail(string sTo,
										 string sToName,
										 string sSubject,
										 string sBody)
		{
			SendMessage(sTo,
						sToName,
						sSubject,
						sBody,
						null);
		}

		public void SendMessage(string sTo,
								string sToName,
								string sSubject,
								string sBody,
								string sFilePath)
		{
			SmtpClient	smtp		= new SmtpClient(m_sSmtp);
			MailAddress	maddrFrom	= new MailAddress(m_sUserId, m_sUserName);
			MailAddress	maddrTo		= new MailAddress(sTo, sToName);
			MailMessage	mmsg		= new MailMessage(maddrFrom, maddrTo);
			string		sPassword	= CSettings.DecryptedPassword(m_sPassword);

			mmsg.SubjectEncoding	= System.Text.Encoding.UTF8;
			mmsg.BodyEncoding		= System.Text.Encoding.UTF8;
			mmsg.Subject			= sSubject;
			mmsg.Body				= sBody;
			if (sFilePath != null) {
				Attachment attachment;
				attachment = new System.Net.Mail.Attachment(sFilePath);
				mmsg.Attachments.Add(attachment);
			}
			smtp.Port				= m_iPort;
			smtp.EnableSsl			= m_bSslTls;
			smtp.Credentials		= new NetworkCredential(m_sUserId, sPassword);
			smtp.Send(mmsg);
			mmsg.Dispose();
			smtp.Dispose();
		}

		public void SendNotificationMailWithAttachment(string sTo,
													   string sToName,
													   string sSubject,
													   string sBody,
													   string sFilePath)
		{
			SendMessage(sTo,
						sToName,
						sSubject,
						sBody,
						sFilePath);
		}
	}
}