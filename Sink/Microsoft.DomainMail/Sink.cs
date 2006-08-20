using System;
using System.DirectoryServices;
using System.Runtime.InteropServices;
using Microsoft.Exchange.Transport.EventInterop;
using Microsoft.Exchange.Transport.EventWrappers;
using System.Reflection;
using System.Threading;
using System.Diagnostics;

namespace Microsoft.DomainMail
{
	[Guid("97B75EC5-C180-4010-8766-C2777EE77F7D")]
	[ComVisible(true)]
	public class Sink : IMailTransportSubmission
	{
		private bool mDebugLog = true;
		private Configuration mConfiguration = null;

		public Sink()
		{
			LoadConfiguration();

			if (DebugLog)
			{
				EventLog.WriteEntry(
					Assembly.GetExecutingAssembly().FullName, 
					"Loaded Microsoft.DomainMail sink.", 
					EventLogEntryType.Information);			
			}
		}

		public bool DebugLog
		{
			get
			{
				return mDebugLog;
			}
		}

		public Configuration Configuration
		{
			get
			{
				LoadConfiguration();
				return mConfiguration;
			}
			set
			{
				mConfiguration = value;
			}
		}

		private void LoadConfiguration()
		{
			Monitor.Enter(this);
			try
			{
				if (mConfiguration == null)
				{
					string ConfigurationFilename = Assembly.GetExecutingAssembly().Location + ".config";
					mConfiguration = new Configuration(ConfigurationFilename);

					object debuglog = mConfiguration["debuglog"];
					mDebugLog = (debuglog == null) ? true : bool.Parse(debuglog.ToString());

					if (DebugLog)
					{
						EventLog.WriteEntry(
							Assembly.GetExecutingAssembly().FullName, 
							"Loaded configuration file " + ConfigurationFilename, 
							EventLogEntryType.Information);
					}
				}
			}
			catch(Exception ex)
			{
				EventLog.WriteEntry(
					Assembly.GetExecutingAssembly().FullName, 
					"Error loading configuration " + Assembly.GetExecutingAssembly().Location + ".config\n" + ex.Message, 
					EventLogEntryType.Error);
			}
			finally
			{
				Monitor.Exit(this);
			}
		}

		public void OnMessageSubmission ( 
			MailMsg message, 
			IMailTransportNotify notify, 
			IntPtr context)
		{
			try
			{
				Message Msg = new Message(message);				

				if (DebugLog)
				{
					EventLog.WriteEntry(
						Assembly.GetExecutingAssembly().FullName, 
						"Checking message " + Msg.Rfc822MsgId + " with subject \"" + Msg.Rfc822MsgSubject + "\".",
						EventLogEntryType.Information);
				}

				RecipsAdd NewRecipients = Msg.AllocNewList();
				bool fReRouted = false;

				foreach (Recip Recipient in Msg.Recips)
				{
					try
					{
						fReRouted |= ReRoute(Recipient, Msg, NewRecipients);
					}
					catch (Exception ex)
					{
						EventLog.WriteEntry(
							Assembly.GetExecutingAssembly().FullName, 
							"Error routing message " + Msg.Rfc822MsgId + " to " + Recipient.SMTPAddress + "." + ex.Message, 
							EventLogEntryType.Error);
					}
				}

				if (fReRouted)
				{
					Msg.WriteList(NewRecipients);
				}

			}
			catch (Exception ex)
			{
				EventLog.WriteEntry(
						Assembly.GetExecutingAssembly().FullName, 
						ex.Message + "\n" + ex.StackTrace.ToString(), 
						EventLogEntryType.Error);
			}
			finally
			{
				if (null != message)
					Marshal.ReleaseComObject(message);
			}

		}

		private bool ReRoute(Recip Recipient, Message Msg, RecipsAdd NewRecipients)
		{
			ActiveDirectory Directory = new ActiveDirectory();

			// TODO: verbose logging
			// Console.WriteLine("Searching for " + proxyAddress + " in " + Directory.UsersLDAPPath.ToString() + ".");

			string[] SearcherPropertiesToLoad = {
				"cn",
				"mail", 
				"proxyAddresses"
			};

			DirectorySearcher Searcher = new DirectorySearcher(
				new DirectoryEntry(Directory.UsersLDAPPath),				
				"(&(objectCategory=person)(objectClass=user)(| (proxyAddresses=*smtp:@" + Recipient.SMTPAddressDomain.ToLower() + "*)(proxyAddresses=*smtp:" + Recipient.SMTPAddress + "*)))",
				SearcherPropertiesToLoad);

			SearchResultCollection SearchResults = Searcher.FindAll();

			if (SearchResults.Count == 0)
				return false;

			foreach (SearchResult SearchResult in SearchResults)
			{
				foreach (string ProxyAddressProperty in SearchResult.Properties["proxyAddresses"])
				{
					string ProxyAddress = ProxyAddressProperty.ToLower();
					if ("smtp:" + Recipient.SMTPAddress.ToLower() == ProxyAddress)
					{
						// there's an address that matches exactly, add him to the re-routing
						// list because there might be other recipients that don't match and
						// will require routing
						NewRecipients.AddSMTPRecipient(Recipient.SMTPAddress, null);
						return false;
					}
				}
			}

			foreach (SearchResult SearchResult in SearchResults)
			{
				foreach (string ProxyAddressProperty in SearchResult.Properties["proxyAddresses"])
				{
					string ProxyAddress = ProxyAddressProperty.ToLower();

					// this is necessary to avoid matching @foo.com with @foo.com.bar
					if ("smtp:@" + Recipient.SMTPAddressDomain.ToLower() == ProxyAddress)
					{
						string RoutedSMTPAddress = SearchResult.Properties["mail"][0].ToString();

						EventLog.WriteEntry(
							Assembly.GetExecutingAssembly().FullName, 
							"Routing message " + Msg.Rfc822MsgId + " from " + Recipient.SMTPAddress + " to " + RoutedSMTPAddress + ".",
							EventLogEntryType.Information);

						NewRecipients.AddSMTPRecipient(RoutedSMTPAddress, null);
						return true;
					}
				}
			}

			return false;	
		}
	}
}
