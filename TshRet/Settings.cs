using System;
using System.Text;
using System.Drawing;
using System.Xml.Serialization;
using System.Reflection;
using System.IO;

namespace TshRet
{
	[Serializable]
	public class CSettings
	{
		public class CContents
		{
			public string	MailSmtp			= "smtp.gmail.com";
			public int		MailSmtpPort		= 587;
			public string	MailImap			= "imap.gmail.com";
			public int		MailImapPort		= 993;
			public bool		MailSslTls			= true;
			public string	MailUserName		= "CAC America Administration";
			public string	MailUserId			= "timesheet@cacamerica.com";
			public string	MailPassword		= "ZkQkdDEkajEkOzEkOXgjblIhZiEh";
			public string	TimeSheetFolder		= string.Empty;
			public string	UploadFileFolder	= string.Empty;

			public CContents()
			{
			}

			~CContents()
			{
			}
		}

		private const string mc_sToken = "harrow";
		private string m_sSettingFile;
		public CContents	Contents;

		public CSettings()
		{
			InitializeSettins();
		}

		~CSettings()
		{
		}

		private void InitializeSettins()
		{
			m_sSettingFile = SettingsDirectory + "\\settings.xml";
			if (File.Exists(m_sSettingFile)) {
				XmlLoad();
				return;
			}
			this.Contents = new CContents();
			Contents.TimeSheetFolder	= Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
			Contents.UploadFileFolder	= Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
			XmlSave();
		}

		public void XmlSave()
		{
			XmlSerializer xmls = new XmlSerializer(typeof(CContents));
			using (FileStream fs = new FileStream(m_sSettingFile, FileMode.Create)) 
			{
				xmls.Serialize(fs, this.Contents);
				fs.Close(); fs.Dispose();
			}
		}

		public void XmlLoad()
		{
			XmlSerializer xmls = new XmlSerializer(typeof(CContents));
			using (FileStream fs = new FileStream(m_sSettingFile, FileMode.Open))
			{
				this.Contents = (CContents)xmls.Deserialize(fs);
				fs.Close(); fs.Dispose();
			}
		}

		public static string SettingsDirectory
		{	
			get {
				Environment.SpecialFolder local = Environment.SpecialFolder.LocalApplicationData;
				string sAppDir = Environment.GetFolderPath(local) + "\\" + ApplicaionTitle;
				if (!Directory.Exists(sAppDir)) Directory.CreateDirectory(sAppDir);
				return sAppDir;
			}
		}

		private static string ApplicaionTitle
		{
			get {
				Assembly asm = Assembly.GetExecutingAssembly();
				return Path.GetFileNameWithoutExtension(asm.GetName().CodeBase);
			}
		}

		public static string DecryptedPassword(string sEncoded)
		{
			byte[] abEncrypted	= Convert.FromBase64String(sEncoded);
			string sEncrypted	= ASCIIEncoding.ASCII.GetString(abEncrypted);
			CCipher cipher		= new CCipher();
			string sPassword;
			cipher.DecodePassword(out sPassword, sEncrypted, mc_sToken);
			return sPassword;
		}

		public static string EncryptedPassword(string sPassword)
		{
			CCipher cipher = new CCipher();
			string sEncrypted;
			cipher.EncodePassword(out sEncrypted, sPassword, mc_sToken);
			byte[] abEncoded	= ASCIIEncoding.ASCII.GetBytes(sEncrypted);
			string sBase64		= Convert.ToBase64String(abEncoded);
			return sBase64;
		}

		public void Delete()
		{
			foreach (string sFile in Directory.GetFiles(SettingsDirectory))
				File.Delete(sFile);
			if (Directory.Exists(SettingsDirectory)) Directory.Delete(SettingsDirectory);
		}
	}
}
