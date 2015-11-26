using System;
using System.Runtime.InteropServices;
using System.Security.Principal;

namespace com.github.yedijas.util
{
    /// <summary>
    /// An utility to ease user impersonation process.
    /// </summary>
    class ImpersonatorUtil
    {
        /// <summary>
        /// Impersonation level to be used.
        /// </summary>
        public enum IMPERSONATION_LEVEL
        {
            Anonymous = 0,
            Identification = 1,
            Impersonation = 2,
            Delegation = 3
        }
        
        /// <summary>
        /// Logon type that are used.
        /// Refer to https://msdn.microsoft.com/en-us/library/windows/desktop/aa378184%28v=vs.85%29.aspx
        /// or http://blogs.msdn.com/b/shawnfa/archive/2005/03/21/400088.aspx for more info.
        /// </summary>
        public enum LOGON_TYPE
        {
            DEFAULT = 0,
            INTERACTIVE = 2,
            NETWORK = 3,
            BATCH = 4,
            SERVICE = 5,
            UNLOCK = 7,
            CLEARTEXT = 8,
            NEW_CREDENTIALS = 9
        }

        /// <summary>
        /// Logon provider that are used.
        /// Refer to https://msdn.microsoft.com/en-us/library/windows/desktop/aa378184%28v=vs.85%29.aspx
        /// for more info.
        /// </summary>
        public enum LOGON_PROVIDER
        {
            DEFAULT = 0,
            WINNT35 = 1,
            WINNT40 = 2,
            WINNT50 = 3
        }

        [DllImport("advapi32.dll", SetLastError = true, CharSet = CharSet.Auto)]
        public static extern int LogonUser(
            string username,
            string domain,
            string password,
            LOGON_TYPE LogonType,
            LOGON_PROVIDER LogonProvider,
            ref IntPtr token
        );

        [DllImport("kernel32.dll", SetLastError = true, CharSet = CharSet.Auto)]
        public static extern bool CloseHandle (IntPtr handle);

        [DllImport("advapi32.dll", SetLastError = true, CharSet = CharSet.Auto)]
        public static extern bool DuplicateToken(
            IntPtr existingTokenHandle, 
            SECURITY_IMPERSONATION_LEVEL impersonationLevel,
            ref IntPtr duplicateTokenHandle
        );

        /// <summary>
        /// Begin impersonation of a specific user using default parameters.
        /// </summary>
        /// <param name="domain">User domain.
        /// Empty and null value will result in current machine</param>
        /// <param name="username">User name to log in.</param>
        /// <param name="password">Password used to log in.</param>
        /// <returns>Impersonated user.</returns>
        public static WindowsImpersonationContext BeginImpersonation(
            string domain,
            string username,
            string password
            )
        {
            WindowsImpersonationContext retVal = BeginImpersonation(
                domain,
                username,
                password,
                LOGON_TYPE.DEFAULT,
                LOGON_PROVIDER.DEFAULT,
                IMPERSONATION_LEVEL.Impersonation);
            return retVal;
        }

        /// <summary>
        /// Begin impersonation of a specific user.
        /// </summary>
        /// <param name="domain">User domain.
        /// Empty and null value will result in current machine</param>
        /// <param name="username">User name to log in.</param>
        /// <param name="password">Password used to log in.</param>
        /// <param name="logonType">Type of logon used.</param>
        /// <param name="logonProvider">Provider used to log in.</param>
        /// <param name="impersonationLevel">Level of impersonation.</param>
        /// <returns>Impersonated user.</returns>
        public static WindowsImpersonationContext BeginImpersonation(
            string domain,
            string username,
            string password,
            LOGON_TYPE logonType,
            LOGON_PROVIDER logonProvider,
            IMPERSONATION_LEVEL impersonationLevel
            )
        {
            IntPtr existingToken = new IntPtr();
            IntPtr duplicateToken = new IntPtr();

            if (string.IsNullOrEmpty(domain))
            {
                domain = System.Environment.MachineName;
            }

            try
            {
                bool isImpersonated = false;
                isImpersonated = LogonUser(
                    username,
                    domain,
                    password,
                    logonType,
                    logonProvider,
                    ref existingToken
                );
                if (isImpersonated)
                {
                    bool isTokenDuplicated = false;
                    isTokenDuplicated = DuplicateToken(
                        existingToken,
                        impersonationLevel,
                        ref duplicateToken);
                    if (isTokenDuplicated)
                    {
                        int errorCode = Marshal.GetLastWin32Error();
                        CloseHandle(existingToken);
                        throw new ApplicationException("Failed to log on with error code : "
                            + errorCode);
                    }
                    else
                    {
                        WindowsIdentity newID = new WindowsIdentity(duplicateToken);
                        WindowsImpersonationContext impersonatedUser = newID.Impersonate();
                        return impersonatedUser;
                    }
                }
                else
                {
                    int errorCode = Marshal.GetLastWin32Error();
                    throw new ApplicationException("Failed to log on with error code : " 
                        + errorCode);
                }
            }
            catch (ApplicationException appEx)
            {
                throw appEx;
            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {
                if (existingToken != IntPtr.Zero)
                {
                    CloseHandle(existingToken);
                }
                if (duplicateToken != IntPtr.Zero)
                {
                    CloseHandle(duplicateToken);
                }
            }
        }

        /// <summary>
        /// Cancels the impersonation.
        /// </summary>
        /// <param name="impersonatedUser">Impersonated user to be undone.</param>
        public void UndoImpersonation(ref WindowsImpersonationContext impersonatedUser)
        {
            impersonatedUser.Undo();
        }
    }
}
