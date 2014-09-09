using System;
using System.Net.Security;
using System.Security.Cryptography.X509Certificates;

namespace OutlookMeetingScraper
{
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