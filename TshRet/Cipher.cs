using System;
using System.Text;
using System.Runtime.InteropServices;

namespace TshRet
{
	class CCipher
	{
		[DllImport("kernel32.dll")]
		private static extern long	RtlMoveMemory(
										byte[]		ayDest,
										string		sSource,
										uint		uiLength
									);

		private string m_sMessage = string.Empty;

		public CCipher()
		{
			return;
		}

		~CCipher()
		{
			return;
		}

		public Boolean DecodePassword(out string sCode, string sCipher, string sKey)
		{
			byte[]	ayHashed;
			byte[]	ayShifted;
			int		iLen;

			sCode = string.Empty;
			try {
				if (sCipher.Length == 0)									return true;
				iLen = CipherStringToHashedArray(out ayHashed, sCipher);
				if (iLen < 0)												return false;
				if (!GatherHashedArray(out ayShifted, ref ayHashed, iLen))	return false;
				if (!UnshiftByteArray(out sCode, ayShifted, sKey, iLen))	return false;
				return true;
			} catch {
				throw new Exception("Decoding the password was failed.");
			}
		}

		private int CipherStringToHashedArray(out byte[] ayHashed, string sCipher)
		{
			int		iLen, iCount, iIndex, iChar, iQuo, iRes, iMax;
			byte[]	ayCipher;

			iLen = (2 * sCipher.Length) / 3;
			iMax = iLen / 2;
			ayHashed = new byte[iLen];
			ayCipher = Encoding.ASCII.GetBytes(sCipher);
			for (iCount = 0; iCount < iMax; iCount++) {
				iIndex = 3 * iCount;
				iChar = ayCipher[iIndex + 2] - 33;
				iChar = 94 * iChar + ayCipher[iIndex + 1] - 33;
				iChar = 94 * iChar + ayCipher[iIndex + 0] - 33;
				iRes = iChar % 256;
				iQuo = (iChar - iRes) / 256;
				ayHashed[2 * iCount] = (byte)iQuo;
				ayHashed[2 * iCount + 1] = (byte)iRes;
			}
			return iLen;
		}

		private Boolean GatherHashedArray(out byte[] ayShifted, ref byte[] ayHashed, int iLen)
		{
		    int		iCount,	iIndex,	iPos;
		    byte[]	ayBit;

		    ayBit	= new byte[iLen * 8];
		    for (iCount = 0; iCount < iLen; iCount++) {
		        for (iIndex = 0; iIndex <= 7; iIndex++) {
		            if ((ayHashed[iCount] & (byte)(1 << iIndex)) == 0)
		                ayBit[(8 * iCount) + iIndex] = 0;
		            else
		                ayBit[(8 * iCount) + iIndex] = 1;
		        }
		    }

		    ayShifted = new byte[iLen];
		    for (iCount = 0; iCount < iLen; iCount++) {
		        ayShifted[iCount] = 0;
		    }
		    iPos = 0;
		    for (iIndex = 0; iIndex <= 7; iIndex++) {
		        for (iCount = 0; iCount < iLen; iCount++) {
		            ayShifted[iCount] = (byte)((int)ayShifted[iCount] + ((int)ayBit[iPos++] << iIndex));
				}
		    }
			return true;
		}

		private Boolean UnshiftByteArray(out string sCode, byte[] ayShifted, string sKey, int iLen)
		{
		    int		iLenK, iWaight, iCount, iShift, iChar, iRes, iQuo;
		    byte[]	ayKey = new byte[1];

		    iLenK = StringToByteArray(ref ayKey, sKey, 1);
		    iWaight = 0;
		    for (iCount = 0; iCount < iLenK; iCount++)
		        iWaight = iWaight + ayKey[iCount];
		    sCode = string.Empty;
		    for (iCount = 0; iCount < iLen; iCount++) {
		        if (iLenK == 0)
		            iShift = 0;
		        else
		            iShift = (8 - ((ayKey[iCount % iLenK] + iWaight) % 8)) % 8;
		        iChar = ayShifted[iCount] * (1 << iShift);
		        iRes = iChar % 256;
		        iQuo = (iChar - iRes) / 256;
			    sCode = sCode + ((char)(iRes + iQuo));
			}
		    if (sCode.Substring(sCode.Length - 1, 1) == "\0") sCode = sCode.Substring(0, sCode.Length - 1);
		    return true;
		}

		private int StringToByteArray(ref byte[] ayArray, string sString, int iBoundary)
		{
			int iLen, iUpper, iCount;

			iLen = sString.Length;
		    if (iLen == 0) return 0;
		    iUpper = BoundaryTableSize(iLen, iBoundary);
		    ayArray = new byte[iUpper];
		    RtlMoveMemory(ayArray, sString, (uint)iLen);
		    for (iCount = iLen; iCount < iUpper; iCount++) {
		        ayArray[iCount] = 0;
			}
		    return iLen;
		}

		public Boolean EncodePassword(out string sCipher, string sCode, string sKey)
		{
			byte[]		ayShifted;
			byte[]		ayHashed;
			int			iLen;

			sCipher = string.Empty;

			try {
				iLen = MakeShiftedByteArray(out ayShifted, sCode, sKey);
				if (iLen < 0)
					throw new Exception("Making shifted table from the password was failed");
				ayHashed = new byte[1];
				if (MakeHashedByteArray(ref ayHashed, ref ayShifted, iLen) == false)
					throw new Exception("Making hashed table from the password was failed");
				MakeCipherString(out sCipher, ayHashed, iLen);
				return true;
			} catch (Exception ex) {
				throw ex;
			}
		}

		private int MakeShiftedByteArray(out byte[] ayShifted, string sCode, string sKey)
		{
			byte[]	ayKey	= new byte[1];
			byte[]	ayCode	= new byte[1];
			int		iLenC, iLenK, iLen;
			int		iWaight, iCount, iShift;
			int		iChar, iRes, iQuo;

			iLenK = StringToByteArray(ref ayKey, sKey, 1);
			iWaight = 0;
			for (iCount = 0; iCount < iLenK; iCount++)
				iWaight += ayKey[iCount];

			iLenC = StringToByteArray(ref ayCode, sCode, 2);
			iLen = BoundaryTableSize(iLenC, 2);
			ayShifted = new byte[iLen];
			for (iCount = 0; iCount < iLenC; iCount++) {
				if (ayCode[iCount] < 0 || ayCode[iCount] > 255) return  -1;
				if (iLenK == 0)
					iShift = 0;
				else
					iShift = (ayKey[iCount % iLenK] + iWaight) % 8;

				iChar = ayCode[iCount] << iShift;
				iRes = iChar % 256;
				iQuo = (iChar - iRes) / 256;
				ayShifted[iCount] = (byte)(iRes + iQuo);
			}
			for (iCount = iLenC; iCount < iLen; iCount++)
				ayShifted[iCount] = 0;

			return iLenC;
		}

		private Boolean MakeHashedByteArray(ref byte[] ayHashed, ref byte[] ayShifted, int iLen)
		{
			int		iLenT, iCount, iIndex, iPos;
			byte[]	ayBit;
		    
			iLenT = BoundaryTableSize(iLen, 2);
			ayBit = new byte[iLenT * 8];
			iPos = 0;
			for (iIndex = 0; iIndex <= 7; iIndex++) {
				for (iCount = 0; iCount < iLenT; iCount++) {
					if (((1 << iIndex) & ayShifted[iCount]) == 0)
						ayBit[iPos] = 0;
					else
						ayBit[iPos] = 1;
					iPos++;
				}
			}
			ayHashed = new byte[iLenT];
			for (iCount = 0; iCount < iLenT; iCount++) {
			   ayHashed[iCount] = 0;
			   for (iIndex = 0; iIndex <= 7; iIndex++)
					  ayHashed[iCount] += (byte)((int)ayBit[(8 * iCount) + iIndex] << iIndex);
			}
			return true;
		}

		private void MakeCipherString(out string sCipher, byte[] ayHashed, int iLen)
		{
			int 			iSize, iMax, iCount, iIndex, iQuo, iRes;
			char[]			acCipher;

			iSize = BoundaryTableSize(iLen, 2);
			acCipher = new char[3 * (iSize / 2)];
			iMax = iSize / 2;
			for (iCount = 0; iCount < iMax; iCount++) {
				iIndex = 3 * iCount;
				iQuo = (int)ayHashed[2 * iCount] * 256 + (int)ayHashed[2 * iCount + 1];
				iRes = iQuo % 94;
				acCipher[iIndex] = (char)(33 + iRes);
				iQuo = (iQuo - iRes) / 94;
				iRes = iQuo % 94;
				acCipher[iIndex + 1]  = (char)(33 + iRes);
				iQuo = (iQuo - iRes) / 94;
				iRes = iQuo % 94;
				acCipher[iIndex + 2] = (char)(33 + iRes);
			}
			sCipher = new string(acCipher);
			return;
		}

		private int BoundaryTableSize(int iLen, int iBoundary)
		{
		    return iLen - ((iLen - 1) % iBoundary) + iBoundary - 1;
		}
	}
}