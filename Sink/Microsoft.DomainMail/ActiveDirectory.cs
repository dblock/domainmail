using System;
using System.Collections;
using System.Threading;
using System.DirectoryServices;

namespace Microsoft.DomainMail
{	
	public class ActiveDirectory
	{

		private DirectoryEntry mRootDSE = null;
		private string mUsersLDAPPath = string.Empty;
		private string mDomainName = string.Empty;
		
		public string DomainName
		{
			get
			{
				return mDomainName;
			}
			set
			{
				mDomainName = value;
				RootDSE = null;
				UsersLDAPPath = string.Empty;
			}
		}

		public string UsersLDAPPath
		{
			get
			{
				Monitor.Enter(this);
				try
				{
					if (mUsersLDAPPath.Length == 0)
					{
						mUsersLDAPPath = "LDAP://" + 
							(DomainName.Length > 0 ? DomainName + "/" : string.Empty) + 
							"CN=Users," + 
							RootDSE.Properties["defaultNamingContext"].Value;
					}
					return mUsersLDAPPath;
				}
				finally
				{
					Monitor.Exit(this);
				}
			}
			set
			{
				mUsersLDAPPath = value;
			}
		}

		public DirectoryEntry RootDSE
		{
			get
			{
				Monitor.Enter(this);
				try
				{
					if (mRootDSE == null)
					{
						mRootDSE = new DirectoryEntry("LDAP://" + 
							(DomainName.Length > 0 ? DomainName + "/" : string.Empty) + 
							"RootDSE");
					}
					return mRootDSE;
				}
				finally
				{
					Monitor.Exit(this);
				}
			}
			set
			{
				mRootDSE = value;
			}
		}

		public ActiveDirectory()
		{

		}


	}
}
