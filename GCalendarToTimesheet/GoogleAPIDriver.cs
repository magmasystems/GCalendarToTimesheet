using System;
using System.Collections.Generic;
using System.IO;
using System.Threading;
using Google.Apis.Auth.OAuth2;
using Google.Apis.Util.Store;

namespace GCalendarToTimesheet
{
    public class GoogleApiDriver
    {
        public UserCredential Authorize(IEnumerable<string> scopes)
        {
            using var stream = new FileStream("credentials.json", FileMode.Open, FileAccess.Read);

            try
            {
                // The file token.json stores the user's access and refresh tokens, and is created
                // automatically when the authorization flow completes for the first time.
                const string credPath = "token.json";
                var credential = GoogleWebAuthorizationBroker.AuthorizeAsync(
                    GoogleClientSecrets.Load(stream).Secrets,
                    scopes,
                    "user",
                    CancellationToken.None,
                    new FileDataStore(credPath, true)).Result;

                return credential;
            }
            catch (Exception e)
            {
                Console.WriteLine(e);
                throw;
            }
        }
    }
}