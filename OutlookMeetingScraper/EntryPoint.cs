using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Security;
using System.Security.Cryptography.X509Certificates;
using Microsoft.Exchange.WebServices.Data;

namespace OutlookMeetingScraper
{
    public class EntryPoint
    {
        public static void Main(string[] args)
        {
            ServicePointManager.ServerCertificateValidationCallback = EwsExample.CertificateValidationCallBack;

            System.Console.WriteLine("Enter your email address");
            var userName = System.Console.ReadLine();

            System.Console.WriteLine("Enter your password.  I won't steal it honest");
            var password = Console.ReadPassword();

            var service = new ExchangeService(ExchangeVersion.Exchange2007_SP1)
            {
                Credentials = new WebCredentials(userName, password),
                TraceEnabled = true,
                TraceFlags = TraceFlags.All
            };

            service.AutodiscoverUrl(userName, EwsExample.RedirectionUrlValidationCallback);          
        }
    }

    /// <summary>
    /// http://stackoverflow.com/questions/3404421/password-masking-console-application
    /// 
    /// Adds some nice help to the console. Static extension methods don't exist (probably for a good reason) so the next best thing is congruent naming.
    /// </summary>
    public static class Console
    {
        /// <summary>
        /// Like System.Console.ReadLine(), only with a mask.
        /// </summary>
        /// <param name="mask">a <c>char</c> representing your choice of console mask</param>
        /// <returns>the string the user typed in </returns>
        private static string ReadPassword(char mask)
        {
            const int ENTER = 13, BACKSP = 8, CTRLBACKSP = 127;
            int[] FILTERED = { 0, 27, 9, 10 /*, 32 space, if you care */ }; // const

            var pass = new Stack<char>();
            char chr = (char)0;

            while ((chr = System.Console.ReadKey(true).KeyChar) != ENTER)
            {
                if (chr == BACKSP)
                {
                    if (pass.Count > 0)
                    {
                        System.Console.Write("\b \b");
                        pass.Pop();
                    }
                }
                else if (chr == CTRLBACKSP)
                {
                    while (pass.Count > 0)
                    {
                        System.Console.Write("\b \b");
                        pass.Pop();
                    }
                }
                else if (FILTERED.Count(x => chr == x) > 0) { }
                else
                {
                    pass.Push((char)chr);
                    System.Console.Write(mask);
                }
            }

            System.Console.WriteLine();

            return new string(pass.Reverse().ToArray());
        }

        /// <summary>
        /// Like System.Console.ReadLine(), only with a mask.
        /// </summary>
        /// <returns>the string the user typed in </returns>
        public static string ReadPassword()
        {
            return ReadPassword('*');
        }
    }

    /// <summary>
    /// All code in this class comes from http://msdn.microsoft.com/en-us/library/office/jj220499(v=exchg.80).aspx
    /// </summary>
    public static class EwsExample
    {
        public static bool RedirectionUrlValidationCallback(string redirectionUrl)
        {
            // The default for the validation callback is to reject the URL.
            bool result = false;

            Uri redirectionUri = new Uri(redirectionUrl);

            // Validate the contents of the redirection URL. In this simple validation
            // callback, the redirection URL is considered valid if it is using HTTPS
            // to encrypt the authentication credentials. 
            if (redirectionUri.Scheme == "https")
            {
                result = true;
            }
            return result;
        }

        public static bool CertificateValidationCallBack(
            object sender,
            X509Certificate certificate,
            X509Chain chain,
            SslPolicyErrors sslPolicyErrors)
        {
            // If the certificate is a valid, signed certificate, return true.
            if (sslPolicyErrors == SslPolicyErrors.None)
            {
                return true;
            }

            // If there are errors in the certificate chain, look at each error to determine the cause.
            if ((sslPolicyErrors & SslPolicyErrors.RemoteCertificateChainErrors) == 0) return false;

            if (chain == null || chain.ChainStatus == null) return true;
            
            foreach (X509ChainStatus status in chain.ChainStatus)
            {
                if (certificate.Subject == certificate.Issuer && status.Status == X509ChainStatusFlags.UntrustedRoot)
                {
                    // Self-signed certificates with an untrusted root are valid. 
                    continue;
                }

                if (status.Status != X509ChainStatusFlags.NoError)
                {
                    // If there are any other errors in the certificate chain, the certificate is invalid,
                    // so the method returns false.
                    return false;
                }
            }

            // When processing reaches this line, the only errors in the certificate chain are 
            // untrusted root errors for self-signed certificates. These certificates are valid
            // for default Exchange server installations, so return true.
            return true;
        }
    }
}
