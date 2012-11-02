/****************************** Module Header ******************************\
File Name:    Default\Ini.cs
Purpose:      General use in any C# project to encrypt data using the AES 
              alogrithm. Also to convert byte-arrays to readable strings and
              the other way around.
By:           Jos van Nijnatten
Date:         29-09-2012
Version:      1.0
\***************************************************************************/

using System;
using System.IO;
using System.Text;
using System.Security.Cryptography;

namespace Default.Utilities
{
    public class Encryption
    {
        #region Properties and variables
        /// <summary>
        /// Key property; Get only.
        /// </summary>
        private byte[] Key;
        public byte[] key { get { return this.Key; } }
        
        /// <summary>
        /// IV property; Get only.
        /// </summary>
        private byte[] IV;
        public byte[] iv { get { return this.IV; } }

        private static Ini encrptIni = new Ini();
        #endregion

        #region Constructor
        /// <summary> Create object </summary>
        public Encryption() : this(new AesManaged()) {
            try
            {
                Debug.write("Setting default Key and IV for encryption");
             
                byte[] _key = null;
                byte[] _iv  = null;

                string __key = encrptIni.ReadString("Encryption", "key", "");
                string __iv  = encrptIni.ReadString("Encryption", "iv",  "");

                AesManaged temp = new AesManaged();

                if ((__key.Split(',').Length == temp.Key.Length) &&
                    (__iv .Split(',').Length == temp.IV .Length))
                {
                    _key = stringToByteArray(__key);
                    _iv  = stringToByteArray(__iv);
                }
                else
                {
                    throw new ArgumentOutOfRangeException("Length of Key or IV in the programs settings file is incorrect");
                }

                this.Key = _key;
                this.IV  = _iv;
            }
            catch (Exception e)
            {
                Debug.write("Error setting default Key and IV for encryption");
                Debug.write(e);
                Debug.write("Creating a new Key and IV for encryption");

                encrptIni.WriteString("Encryption", "key", byteArrayToReadableString(this.Key));
                encrptIni.WriteString("Encryption", "iv",  byteArrayToReadableString(this.IV));
            }
        }
        public Encryption(AesManaged myCrypt) : this(myCrypt.Key, myCrypt.IV) { }
        public Encryption(byte[] Key, byte[] IV) 
        {
            AesManaged temp = new AesManaged();
            if (Key == null)
            {
                throw new NullReferenceException("Key");
            }
            else if (IV == null)
            {
                throw new NullReferenceException("IV");
            }
            else if (Key.Length != temp.Key.Length)
            {
                throw new ArgumentOutOfRangeException("Length of Key is incorrect");
            }
            else if (IV.Length != temp.IV.Length)
            {
                throw new ArgumentOutOfRangeException("Length of IV is incorrect");
            }
            else
            {
                this.Key = Key;
                this.IV  = IV;
            }
        }
        #endregion


        #region MD5 Hash
        /// <summary>
        /// Creates a MD5 hash out of a string
        /// </summary>
        /// <param name="input"></param>
        /// <returns></returns>
        public static string md5(string input)
        {
            MD5CryptoServiceProvider crypt = new MD5CryptoServiceProvider();
            byte[] bs = crypt.ComputeHash(stringToByteArray(input));
            return byteArrayToString(bs);
        }
        #endregion

        #region string <-> byte[] (NO Encryption!)
        /// <summary>
        /// Converts a string to a byte array
        /// </summary>
        /// <param name="input"></param>
        /// <returns></returns>
        public static byte[] stringToByteArray(string input)
        {
            if (!string.IsNullOrEmpty(input))
            {
                return Encoding.ASCII.GetBytes(input); //output;
            }
            #region error
            else
            {
                throw new NullReferenceException("'String input' can't be NULL or empty");
            }
            #endregion
        }

        /// <summary>
        /// Converts a byte array to a string
        /// </summary>
        /// <param name="input"></param>
        /// <returns></returns>
        public static string byteArrayToString(byte[] input)
        {
            if (input != null)
            {
                return ASCIIEncoding.ASCII.GetString(input);
            }
            #region error
            else
            {
                throw new NullReferenceException("'byte[] input' can't be NULL");
            }
            #endregion
        }
        #endregion

        #region string <-> byte[] (AES Encryption, with Key and IV)
        /// <summary>
        /// Encrype a string to bytes
        /// </summary>
        /// <param name="plainText"></param>
        /// <returns></returns>
        public byte[] encryptStringToBytes(string plainText)
        {
            // Check arguments.
            if (plainText == null || plainText.Length <= 0)
                throw new ArgumentNullException("plainText");
            if (Key == null || Key.Length <= 0)
                throw new ArgumentNullException("Key");
            if (IV == null || IV.Length <= 0)
                throw new ArgumentNullException("Key");

            // Declare the streams used
            // to encrypt to an in memory
            // array of bytes.
            MemoryStream msEncrypt = null;
            CryptoStream csEncrypt = null;
            StreamWriter swEncrypt = null;

            // Declare the Managed object
            // used to encrypt the data.
            AesManaged alg = null;

            try
            {
                // Create a Managed object
                // with the specified key and IV.
                alg = new AesManaged();
                alg.Key = this.Key;
                alg.IV  = this.IV;
                    
                // Create a encryptor to perform the stream transform.
                ICryptoTransform encryptor = alg.CreateEncryptor();

                // Create the streams used for encryption.
                msEncrypt = new MemoryStream();
                csEncrypt = new CryptoStream(msEncrypt, encryptor, CryptoStreamMode.Write);
                swEncrypt = new StreamWriter(csEncrypt);

                //Write all data to the stream.
                swEncrypt.Write(plainText);
                
            }
            finally
            {
                // Clean things up.

                // Close the streams.
                if (swEncrypt != null)
                    swEncrypt.Close();
                if (csEncrypt != null)
                    csEncrypt.Close();
                if (msEncrypt != null)
                    msEncrypt.Close();

                // Clear the Managed object.
                if (alg != null)
                {
                    alg.Clear();
                }
            }

            // Return the encrypted bytes from the memory stream.
            return msEncrypt.ToArray();
        }

        /// <summary>
        /// Decrype bytes to string
        /// </summary>
        /// <param name="cipherText"></param>
        /// <returns></returns>
        public string decrypeBytesToString(byte[] cipherText)
        {

            // Check arguments.
            if (cipherText == null || cipherText.Length <= 0)
                throw new ArgumentNullException("cipherText");
            if (Key == null || Key.Length <= 0)
                throw new ArgumentNullException("Key");
            if (IV == null || IV.Length <= 0)
                throw new ArgumentNullException("Key");

            // TDeclare the streams used
            // to decrypt to an in memory
            // array of bytes.
            MemoryStream msDecrypt = null;
            CryptoStream csDecrypt = null;
            StreamReader srDecrypt = null;

            // Declare the Managed object
            // used to decrypt the data.
            AesManaged alg = null;

            // Declare the string used to hold
            // the decrypted text.
            string plaintext = null;

            try
            {
                // Create a Managed object
                // with the specified key and IV.
                alg = new AesManaged();
                alg.Key = this.Key;
                alg.IV  = this.IV;
                
                // Create a decrytor to perform the stream transform.
                ICryptoTransform decryptor = alg.CreateDecryptor();

                // Create the streams used for decryption.
                msDecrypt = new MemoryStream(cipherText);
                csDecrypt = new CryptoStream(msDecrypt, decryptor, CryptoStreamMode.Read);
                srDecrypt = new StreamReader(csDecrypt);

                // Read the decrypted bytes from the decrypting stream
                // and place them in a string.
                plaintext = srDecrypt.ReadToEnd();
                
            }
            finally
            {
                // Clean things up.

                // Close the streams.
                if (srDecrypt != null)
                    srDecrypt.Close();
                if (csDecrypt != null)
                    csDecrypt.Close();
                if (msDecrypt != null)
                    msDecrypt.Close();

                // Clear the Managed object.
                if (alg != null)
                {
                    alg.Clear();
                }
            }

            return plaintext;
        }
        #endregion

        #region byte <-> readable string (NO Encryption!)
        /// <summary>
        /// Convert a byte-array to a string in the readable spectrum
        /// </summary>
        /// <param name="input">the byte-array</param>
        /// <returns>Value of each byte, separated by a comma</returns>
        public static string byteArrayToReadableString(byte[] input)
        {
            #region Errors
            if (input == null)
            {
                throw new NullReferenceException("'byte[] Input' can't be NULL");
            }
            #endregion

            StringBuilder output = new StringBuilder();

            for (int i = 0; i < input.Length; i++)
            {
                output.Append((int)(input[i]));
                if (i != (input.Length-1))
                {
                    output.Append(',');
                }
            }

            return output.ToString();
        }

        /// <summary>
        /// Convert a string (e.g. "42,56,76,2,64") to a byte array
        /// </summary>
        /// <param name="input"></param>
        /// <returns></returns>
        public static byte[] readableStringToByteArray(string input)
        {
            #region Errors
            if (string.IsNullOrEmpty(input))
            {
                throw new NullReferenceException("'String input' can't be NULL or empty");
            }
            #endregion

            // Try to convert numbers to bytes
            string[] numbers = input.Split(',');
            byte[] output = new byte[numbers.Length];

            for (int i = 0; i < numbers.Length; i++)
            {
                int res = -1;
                if ((!string.IsNullOrEmpty(numbers[i])) && (int.TryParse(numbers[i], out res)))
                {
                    output[i] = byte.Parse(numbers[i]);
                }
                #region Errors
                else
                {
                    Debug.write("Could not convert string ('" + input + "') to byte-array.");
                    if (numbers[i] == null)
                    {
                        throw new NullReferenceException("Item #" + i + " from input object is null");
                    }
                    else if (numbers[i] == "")
                    {
                        throw new Exception("Item #" + i + " from input object is empty");
                    }
                    else if (res < 0)
                    {
                        throw new Exception("Item #" + i + " from input object is not numberic");
                    }
                    else
                    {
                        throw new Exception("Unknown error at item #" + i + " from input object");
                    }
                }
                #endregion
            }

            return output;
        }
        #endregion
    }
}
