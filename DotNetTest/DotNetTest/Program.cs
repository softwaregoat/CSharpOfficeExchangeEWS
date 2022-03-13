using Microsoft.Exchange.WebServices.Data;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DotNetTest
{
    class Program
    {
        static void Main(string[] args)
        {
            var user_email = ConfigurationManager.AppSettings.Get("user_email");
            var user_password = ConfigurationManager.AppSettings.Get("user_password");

            ExchangeService _service;

            try
            {
                Console.WriteLine("Registering Exchange connection");

                _service = new ExchangeService
                {
                    Credentials = new WebCredentials(user_email, user_password)
                };
            }
            catch
            {
                Console.WriteLine("new ExchangeService failed. Press enter to exit:");
                return;
            }

            // This is the office365 webservice URL
            _service.Url = new Uri("https://outlook.office365.com/EWS/Exchange.asmx");

            // Prepare seperate class for writing email to the database
            try
            {

                Console.WriteLine("Reading mail");

                var path = "Result.txt";
                TextWriter tw = new StreamWriter(path);

                // Read 100 items
                foreach (Contact contact in _service.FindItems(WellKnownFolderName.Contacts, new ItemView(100)))
                {
                    var GivenName = contact.GivenName;
                    var Surname = contact.Surname;
                    var City = contact.PhysicalAddresses[PhysicalAddressKey.Business].City;
                    var CountryOrRegion = (contact.PhysicalAddresses[PhysicalAddressKey.Business].CountryOrRegion);
                    var PostalCode = (contact.PhysicalAddresses[PhysicalAddressKey.Business].PostalCode);
                    var State = (contact.PhysicalAddresses[PhysicalAddressKey.Business].State);
                    var Street = (contact.PhysicalAddresses[PhysicalAddressKey.Business].Street);
                    var MobilePhone = (contact.PhoneNumbers[PhoneNumberKey.MobilePhone]);
                    var EmailAddress1 = contact.EmailAddresses[EmailAddressKey.EmailAddress1];

                    var line = String.Format("{0}|{1}|{2} {3}, {4}, {5}|{6}|{7}|{8}",
                        GivenName, Surname, City, CountryOrRegion, PostalCode, State, Street, MobilePhone, EmailAddress1);

                    tw.WriteLine(line);

                    Console.WriteLine(line);
                }

                tw.Close();
                Console.WriteLine("Exiting");
            }
            catch (Exception e)
            {
                Console.WriteLine("An error has occured. \n:" + e.Message);
            }
            Console.ReadLine();
        }
    }
}
