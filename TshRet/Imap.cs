using System;
using System.Text;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Net.Sockets;
using System.Net.Security;
using System.Configuration;

namespace TshRet
{
	public class CImap
	{
		private CSettings m_settings;

		public CImap(CSettings settings)
		{
			m_settings = settings;
		}

		~CImap()
		{
		}

		public string DonwloadTimeSheetMail()
		{
			string	sUserid, sPassword;
			GetMailUseridAndPassword(out sUserid, out sPassword);

			string sTempDir = GetTempDir();
			string sPrefix	= GetTempFileNamePrefix();
			string sLpath = sTempDir + "\\" + sPrefix + "maillog.txt";
//			if (File.Exists(sLpath)) File.Delete(sLpath);
			string sDpath = sTempDir  + "\\" + sPrefix + "mailtext.txt";
//			if (File.Exists(sDpath)) File.Delete(sDpath);
			string sEpath = sTempDir  + "\\" + sPrefix + "timesheet";
			string sRes;

			StreamWriter	strmlgw = new System.IO.StreamWriter(System.IO.File.Create(sLpath));
			TcpClient		tcpclnt	= new System.Net.Sockets.TcpClient("imap.gmail.com", 993);
			SslStream		sslstrm	= new System.Net.Security.SslStream(tcpclnt.GetStream());
	
			// There should be no gap between the imap command and the \r\n 
			// sslstrm.read() -- while sslstrm.readbyte!= eof does not work 
			// because there is no eof from server and  cannot check for \r\n  
			// because in case of larger response from server ex:read email 
			// message. There are lot of lines so \r\n appears at the end of 
			// each line sslstrm.timeout sets the underlying tcp connections 
			// timeout if the read or writetime out exceeds then the undelying 
			// connectionis closed.

			strmlgw.WriteLine("### START ====================================", sslstrm);
			sslstrm.AuthenticateAsClient("imap.gmail.com");
			strmlgw.WriteLine("# Send blank =================================", sslstrm);
			SendImapCommand("", sslstrm);
			sRes = ReceiveImapRespons(sslstrm);
			strmlgw.Write(sRes);

			strmlgw.WriteLine("# Send 'LOGIN' ===============================", sslstrm);
			SendImapCommand("$ LOGIN " + sUserid + " " + sPassword + "  \r\n", sslstrm);
			sRes = ReceiveImapRespons(sslstrm);
			strmlgw.Write(sRes);

			strmlgw.WriteLine("# Send 'SELECT INBOX' ========================", sslstrm);
			SendImapCommand("$ SELECT INBOX\r\n", sslstrm);
			sRes = ReceiveImapRespons(sslstrm);
			strmlgw.Write(sRes);

//			strmlgw.WriteLine("# Send 'TATUS INBOX (MESSAGES)' ==============", sslstrm);
//			SendImapCommand("$ STATUS INBOX (MESSAGES)\r\n", sslstrm);
//			sRes = ReceiveImapRespons(sslstrm);
//			strmlgw.Write(sRes);
//			int number = 1;
//			strmlgw.WriteLine("# Send 'FETCH " + number.ToString() + " BODYSTRUCTURE' =======", sslstrm);
//			SendImapCommand("$ FETCH " + number + " bodystructure\r\n", sslstrm);
//			sRes = ReceiveImapRespons(sslstrm);
//			strmlgw.Write(sRes);
//			strmlgw.WriteLine("# Send 'FETCH " + number.ToString() + " BODY[HEADER]' ========", sslstrm);
//			SendImapCommand("$ FETCH " + number + " body[header]\r\n", sslstrm);
//			sRes = ReceiveImapRespons(sslstrm);
//			strmlgw.Write(sRes);

			strmlgw.WriteLine("# Send 'FETCH 1 body[text]' ==================", sslstrm);
			SendImapCommand("$ FETCH 1 body[text]\r\n", sslstrm);
			ReceiveMailBody(sDpath, tcpclnt, sslstrm);

			string sTempTimeSheetPath = ExtractAttachedFile(sDpath, sEpath);

			strmlgw.WriteLine("# Send 'LOGOUT' ==============================", sslstrm);
			SendImapCommand("$ LOGOUT\r\n", sslstrm);
			sRes = ReceiveImapRespons(sslstrm);
			strmlgw.Write(sRes);
			strmlgw.WriteLine("### END ======================================", sslstrm);

			strmlgw.Close(); strmlgw.Dispose();
			sslstrm.Close(); sslstrm.Dispose();
			tcpclnt.Close();

			File.Delete(sLpath);
			File.Delete(sDpath);

			return sTempTimeSheetPath;
		}

		private string GetTempDir()
		{
//			Environment.SpecialFolder es = Environment.SpecialFolder.LocalApplicationData;
//			string sTempDir = Environment.GetFolderPath(es) + "\\Temp";
			return  "C:\\Temp";
		}

		private string GetTempFileNamePrefix()
		{
			DateTime dt = DateTime.Now;
			Random rand = new Random();
			int iRand = (int)(rand.NextDouble() * 10000);
			string sFileTitle = "temp" + dt.Second.ToString("00") + dt.Millisecond.ToString("000");
			return sFileTitle;
		}

		private void GetMailUseridAndPassword(out string sUserid, out string sPassword)
		{
			sUserid				= m_settings.Contents.MailUserId;
			string sEncripted	= m_settings.Contents.MailPassword;
			sPassword			= CSettings.DecryptedPassword(sEncripted);
		}

		private void SendImapCommand(string			sCommand,
									 SslStream		sslstrm)
		{
			try {
				byte[]	abCommand;
				abCommand = Encoding.ASCII.GetBytes(sCommand);
				sslstrm.Write(abCommand, 0, abCommand.Length);
			} catch (Exception ex) {
				throw new ApplicationException(ex.Message);
			}
		}

		private string ReceiveImapRespons(SslStream sslstrm)
		{
			try {
				StringBuilder strb = new StringBuilder();
				byte[]	abBuffer;
				sslstrm.Flush();
				int iBuffLen = 4096;
				abBuffer	= new byte[iBuffLen];
				int iRead	= 0;
				while (true) {
					iRead = sslstrm.Read(abBuffer, 0, iBuffLen);
					int iTail;
					for (iTail = 0; iTail < iRead; iTail++)
						if (abBuffer[iTail] == 0) break;
					string sBuff = Encoding.ASCII.GetString(abBuffer, 0, iTail);
					strb.Append(sBuff);
					if (iTail < iBuffLen) break;
				}
				string sResponce = strb.ToString();
				return sResponce;
			} catch (Exception ex) {
				throw new ApplicationException(ex.Message);
			}

		}

		private void ReceiveMailBody(string			sDpath,
									 TcpClient		tcpclnt,
									 SslStream		sslstrm)
		{
			StreamWriter strmdtw	= new System.IO.StreamWriter(System.IO.File.Create(sDpath));
	
			StringBuilder streamb	= new StringBuilder();
			sslstrm.Flush();

			string	sLine			= ReadFirstLine(sslstrm);
			int		iDataLen		= GetBodyStreamLength(sLine);
			if (iDataLen < 0)
				throw new Exception("Unrecognized format of mail text first line:\n"
									+ "\"" + sLine + "\"");
			int		iBufLen			= 4096;
			byte[]	abBuffer;
			int		iBytes;
			abBuffer = new byte[iBufLen];
			iBytes = 0;
			int iRead, iRem;
			iRem = iDataLen - iBytes;
			while (iRem > 0) {
				iRead = (iRem > iBufLen)?iBufLen:iRem;
				iRead =  sslstrm.Read(abBuffer, 0, iRead);
				streamb.Append(Encoding.ASCII.GetString(abBuffer, 0, iRead));
				iBytes += iRead;
				iRem = iDataLen - iBytes;
			}
			strmdtw.Write(streamb.ToString());
			strmdtw.Close(); strmdtw.Dispose();
		}

		private string ReadFirstLine(SslStream sslstrm)
		{
			char[] acBuffer = new char[256];
			int iInput	= sslstrm.ReadByte();
			int iByte	= 0;
			while (iInput >= 0) {
				if (iInput == 0x0d) break;
				if (iByte >= 256)
					throw new Exception("Unrecognized format of mail text first line.");
				acBuffer[iByte++] = (char)iInput;
				iInput = sslstrm.ReadByte();
			}
			acBuffer[iByte++] = (char)iInput;
			iInput = sslstrm.ReadByte();
			acBuffer[iByte++] = (char)iInput;
			char[] acResult = new char[iByte];
			for (int i = 0; i < iByte; i++)
				acResult[i] = acBuffer[i];
			return new string(acResult);
		}

		private int GetBodyStreamLength(string sFirstLine)
		{
			int iPosS = sFirstLine.LastIndexOf('{') + 1;
			if (iPosS < 0) return -1;
			int iPosE = sFirstLine.LastIndexOf('}') - 1;
			if (iPosE < 0) return -1;
			if (iPosE <= iPosS) return -1;
			string sLen = sFirstLine.Substring(iPosS, iPosE - iPosS + 1);
			int iLen;
			bool bState = int.TryParse(sLen, out iLen);
			if (!bState) return -1;
			return iLen;
		}

		private string ExtractAttachedFile(string sDpath, string sEpath)
		{
			TextReader txtrDt = new StreamReader(sDpath);
			string sSeparator  = txtrDt.ReadLine();
			string sLine = txtrDt.ReadLine();
			while (sLine != sSeparator) sLine = txtrDt.ReadLine();
			sLine = txtrDt.ReadLine();
			string sFileExtention = GetAttachedFilenameExtention(sLine);
			while (sLine.Length > 0) sLine = txtrDt.ReadLine();
			StringBuilder sbData = new StringBuilder();
			sbData.Clear();
			sLine = txtrDt.ReadLine();
			while (sLine.Substring(0, sSeparator.Length) != sSeparator) {
				if (txtrDt.Peek() < 0) break;
				sbData.Append(sLine);
				sLine = txtrDt.ReadLine();
			}
			string sXlsPath = sEpath + sFileExtention;
			BinaryWriter bnrywEx = new BinaryWriter(File.Open(sXlsPath, FileMode.Create, FileAccess.Write));
			byte[] abXlsx = Convert.FromBase64String(sbData.ToString());
			bnrywEx.Write(abXlsx);
			bnrywEx.Close(); bnrywEx.Dispose();
			txtrDt.Close(); txtrDt.Dispose();

			return sXlsPath;
		}

		private string GetAttachedFilenameExtention(string sLine)
		{
			int iPeriod = sLine.LastIndexOf('.');
			string sExtention = sLine.Substring(iPeriod, sLine.Length - iPeriod - 1);
			return sExtention;
		}
	}
}