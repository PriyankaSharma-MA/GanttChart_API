using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Text;
using System.Security.Cryptography;
using System.IO;

namespace GanttChart.App_Code
{
    public class Decryption
    {
        public static string Decrypt(string input)
        {
            byte[] resultArray = null;
            string result = input;
            try
            {
                byte[] inputArray = Convert.FromBase64String(input);
                TripleDESCryptoServiceProvider tripleDES = new TripleDESCryptoServiceProvider();
                tripleDES.Key = UTF8Encoding.UTF8.GetBytes("sblw-3hn8-sqoy19");
                tripleDES.Mode = CipherMode.ECB;
                tripleDES.Padding = PaddingMode.PKCS7;
                ICryptoTransform cTransform = tripleDES.CreateDecryptor();
                resultArray = cTransform.TransformFinalBlock(inputArray, 0, inputArray.Length);
                result = "";
                result = UTF8Encoding.UTF8.GetString(resultArray);
                tripleDES.Clear();
            }

            catch (Exception ex)
            {
                result = input;
            }
            return result;
        }

        public static string DecryptNew(string cryptTxt)
        {
            //bool isEncrypt = false;
            //try
            //{
            //    byte[] bytesBuff = Convert.FromBase64String(cryptTxt);
            //    isEncrypt = true;
            //}
            //catch (Exception ex)
            //{
            //    isEncrypt = false;
            //}


            // if (isEncrypt)
            if (!string.IsNullOrWhiteSpace(cryptTxt))
            {
                try
                {
                    string key = "encryption";
                    // cryptTxt = cryptTxt.Replace(" ", "+");
                    //byte[] bytesBuff = Convert.FromBase64String(cryptTxt);
                    //using (Aes aes = Aes.Create())
                    //{
                    //    Rfc2898DeriveBytes crypto = new Rfc2898DeriveBytes(key,
                    //        new byte[] { 0x49, 0x76, 0x61, 0x6e, 0x20, 0x4d, 0x65, 0x64, 0x76, 0x65, 0x64, 0x65, 0x76 });
                    //    aes.Key = crypto.GetBytes(32);
                    //    aes.IV = crypto.GetBytes(16);
                    //    using (MemoryStream mStream = new MemoryStream())
                    //    {
                    //        using (CryptoStream cStream = new CryptoStream(mStream, aes.CreateDecryptor(), CryptoStreamMode.Write))
                    //        {
                    //            cStream.Write(bytesBuff, 0, bytesBuff.Length);
                    //            cStream.Close();
                    //        }
                    //        cryptTxt = Encoding.Unicode.GetString(mStream.ToArray());
                    //    }
                    //}


                    string[] encodedTextArray = cryptTxt.Split(new string[] { "||" }, StringSplitOptions.None);
                    string decodedText = string.Empty;
                    foreach (string str in encodedTextArray)
                    {
                        decodedText += Convert.ToChar(Convert.ToInt32(str) + 2);//+2 as we have minus the value while encoding
                    }
                    cryptTxt = decodedText;
                }
                catch (Exception ex)
                {
                    
                    string filePath = Path.Combine(HttpContext.Current.Server.MapPath("~/Log"), "Error.txt");
                    using (StreamWriter writer = new StreamWriter(filePath, true))
                    {
                        writer.WriteLine("-----------------------------------------------------------------------------");
                        writer.WriteLine("Date : " + DateTime.Now.ToString());
                        writer.WriteLine();

                        while (ex != null)
                        {
                            writer.WriteLine(ex.GetType().FullName);
                            writer.WriteLine("Message : " + ex.Message);
                            writer.WriteLine("StackTrace : " + ex.StackTrace);

                            ex = ex.InnerException;
                        }
                    }
                }
            }
            return cryptTxt;
        }

        public void writeLog(Exception ex)
        {
            string filePath = Path.Combine(HttpContext.Current.Server.MapPath("~/Log"), "Error.txt");
            using (StreamWriter writer = new StreamWriter(filePath, true))
            {
                writer.WriteLine("-----------------------------------------------------------------------------");
                writer.WriteLine("Date : " + DateTime.Now.ToString());
                writer.WriteLine();

                while (ex != null)
                {
                    writer.WriteLine(ex.GetType().FullName);
                    writer.WriteLine("Message : " + ex.Message);
                    writer.WriteLine("StackTrace : " + ex.StackTrace);

                    ex = ex.InnerException;
                }
            }
        }
        public static string DecryptNew_old(string cryptTxt)
        {
            bool isEncrypt = false;
            try
            {
                byte[] bytesBuff = Convert.FromBase64String(cryptTxt);
                isEncrypt = true;
            }
            catch (Exception ex)
            {
                isEncrypt = false;
            }
            if (isEncrypt)
            {
                string key = "encryption";
                cryptTxt = cryptTxt.Replace(" ", "+");
                byte[] bytesBuff = Convert.FromBase64String(cryptTxt);
                using (Aes aes = Aes.Create())
                {
                    Rfc2898DeriveBytes crypto = new Rfc2898DeriveBytes(key,
                        new byte[] { 0x49, 0x76, 0x61, 0x6e, 0x20, 0x4d, 0x65, 0x64, 0x76, 0x65, 0x64, 0x65, 0x76 });
                    aes.Key = crypto.GetBytes(32);
                    aes.IV = crypto.GetBytes(16);
                    using (MemoryStream mStream = new MemoryStream())
                    {
                        using (CryptoStream cStream = new CryptoStream(mStream, aes.CreateDecryptor(), CryptoStreamMode.Write))
                        {
                            cStream.Write(bytesBuff, 0, bytesBuff.Length);
                            cStream.Close();
                        }
                        cryptTxt = Encoding.Unicode.GetString(mStream.ToArray());
                    }
                }
            }
            return cryptTxt;
        }
    }
}