using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Security.Cryptography;

namespace KalevaAalto
{
    /// <summary>
    /// 加密技术
    /// </summary>
    public static partial class Main
    {
        /// <summary>
        /// SHA1加密工具
        /// </summary>
        private static SHA1 sha1 = SHA1.Create();
        /// <summary>
        /// 对字符串进行SHA1加密
        /// </summary>
        /// <param name="input">要加密的密文</param>
        /// <returns></returns>
        public static string ComputeSha1Hash(this string input)
        {
            byte[] inputBytes = Encoding.UTF8.GetBytes(input);
            byte[] hashBytes = sha1.ComputeHash(inputBytes);

            return Convert.ToBase64String(hashBytes);
        }

        /// <summary>
        /// 获取随机盐
        /// </summary>
        public static string salt
        {
            get
            {
                return GetRandomString(20,30);
            }
        }


        // 使用AES密钥和初始化向量（IV）加密数据
        public static byte[] EncryptData(byte[] data, byte[] key, byte[] iv)
        {
            using (Aes aes = Aes.Create())
            {
                aes.Key = key;
                aes.IV = iv;

                using (ICryptoTransform encryptor = aes.CreateEncryptor())
                {
                    return PerformCryptography(data, encryptor);
                }
            }
        }

        // 使用AES密钥和初始化向量（IV）解密数据
        public static byte[] DecryptData(byte[] encryptedData, byte[] key, byte[] iv)
        {
            using (Aes aes = Aes.Create())
            {
                aes.Key = key;
                aes.IV = iv;

                using (ICryptoTransform decryptor = aes.CreateDecryptor())
                {
                    return PerformCryptography(encryptedData, decryptor);
                }
            }
        }
        // 执行加密或解密操作
        private static byte[] PerformCryptography(byte[] data, ICryptoTransform cryptoTransform)
        {
            using (MemoryStream ms = new MemoryStream())
            {
                using (CryptoStream cs = new CryptoStream(ms, cryptoTransform, CryptoStreamMode.Write))
                {
                    cs.Write(data, 0, data.Length);
                    cs.FlushFinalBlock();
                    return ms.ToArray();
                }
            }
        }
    }
}
