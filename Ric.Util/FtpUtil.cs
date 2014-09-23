using System;
using System.Diagnostics;
using System.IO;
using System.Net;
using System.Net.Sockets;
using System.Text;

namespace Ric.Util
{
    /// <summary>
    /// FTP related operaions: Download file /Upload file
    /// </summary>
    public class FtpUtil
    {
        /// <summary>
        /// Anonymously download file from FTP service
        /// </summary>
        /// <param name="remoteFileUri"></param>
        /// <param name="localFilePath"></param>
        public static void DownloadFile(string remoteFileUri, string localFilePath)
        {
            Uri uri = new Uri(remoteFileUri);
            string[] userInfo = string.IsNullOrEmpty(uri.UserInfo) ? new[] { "anonymous", "" } : uri.UserInfo.Split(':');
            string host = uri.Host;
            string userName = userInfo[0];
            string password = (userInfo.Length > 1 ? userInfo[1] : "");
            string remoteFileDir = Path.GetDirectoryName(uri.PathAndQuery);
            string remoteFileName = Path.GetFileName(uri.PathAndQuery);

            FtpClient fc = new FtpClient(host, userName, password);
            fc.Login();
            fc.ChangeDir(remoteFileDir);
            fc.Download(remoteFileName, Path.GetFullPath(localFilePath));
            fc.Close();
        }

        /// <summary>
        /// Anonymously upload file to FTP service
        /// </summary>
        /// <param name="remoteDirUri"></param>
        /// <param name="localFilePath"></param>
        public static void UploadFile(string remoteDirUri, string localFilePath)
        {
            Uri uri = new Uri(remoteDirUri);
            string[] userInfo = string.IsNullOrEmpty(uri.UserInfo) ? new[] { "anonymous", "" } : uri.UserInfo.Split(':');
            string host = uri.Host;
            string userName = userInfo[0];
            string password = (userInfo.Length > 1 ? userInfo[1] : "");
            string remoteFileDir = uri.PathAndQuery;

            FtpClient fc = new FtpClient(host, userName, password);
            fc.Login();
            fc.ChangeDir(remoteFileDir);
            fc.Upload(localFilePath);
            fc.Close();
        }
    }

    /// <summary>
    /// FTP Client Related
    /// </summary>
    public class FtpClient
    {
        public class FtpException : Exception
        {
            public FtpException(string message)
                : base(message)
            {
            }

            public FtpException(string message, Exception innerException)
                : base(message, innerException)
            {
            }
        }

        private const int BufferSize = 512;
        private static readonly Encoding Ascii = Encoding.ASCII;

        private bool _verboseDebugging;

        // defaults
        private string _server = "localhost";

        private string _remotePath = ".";
        private string _username = "anonymous";
        private string _password = "anonymous@anonymous.net";
        private string _message;
        private string _result;

        private int _port = 21;
        private int _bytes;
        private int _resultCode;

        private bool _loggedin;
        private const bool BinMode = false;

        private readonly Byte[] _buffer = new Byte[BufferSize];
        private Socket _clientSocket;

        private int _timeoutSeconds = 10;

        /// <summary>
        /// Default contructor
        /// </summary>
        public FtpClient()
        {
        }

        /// <summary>
        ///
        /// </summary>
        /// <param name="server"></param>
        /// <param name="username"></param>
        /// <param name="password"></param>
        public FtpClient(string server, string username, string password)
        {
            _server = server;
            _username = username;
            _password = password;
        }

        /// <summary>
        /// Create a new FTP client
        /// </summary>
        /// <param name="server">FTP server</param>
        /// <param name="username"></param>
        /// <param name="password"></param>
        /// <param name="timeoutSeconds"></param>
        /// <param name="port"></param>
        public FtpClient(string server, string username, string password, int timeoutSeconds, int port)
        {
            _server = server;
            _username = username;
            _password = password;
            _timeoutSeconds = timeoutSeconds;
            _port = port;
        }

        /// <summary>
        /// Display all communications to the debug log
        /// </summary>
        public bool VerboseDebugging
        {
            get
            {
                return _verboseDebugging;
            }
            set
            {
                _verboseDebugging = value;
            }
        }

        /// <summary>
        /// Remote server port. Typically TCP 21
        /// </summary>
        public int Port
        {
            get
            {
                return _port;
            }
            set
            {
                _port = value;
            }
        }

        /// <summary>
        /// Timeout waiting for a response from server, in seconds.
        /// </summary>
        public int Timeout
        {
            get
            {
                return _timeoutSeconds;
            }
            set
            {
                _timeoutSeconds = value;
            }
        }

        /// <summary>
        /// Gets and Sets the name of the FTP server.
        /// </summary>
        /// <returns></returns>
        public string Server
        {
            get
            {
                return _server;
            }
            set
            {
                _server = value;
            }
        }

        /// <summary>
        /// Gets and Sets the port number.
        /// </summary>
        /// <returns></returns>
        public int RemotePort
        {
            get
            {
                return _port;
            }
            set
            {
                _port = value;
            }
        }

        /// <summary>
        /// Gets and Sets the remote directory.
        /// </summary>
        public string RemotePath
        {
            get
            {
                return _remotePath;
            }
            set
            {
                _remotePath = value;
            }
        }

        /// <summary>
        /// Gets and Sets the username.
        /// </summary>
        public string Username
        {
            get
            {
                return _username;
            }
            set
            {
                _username = value;
            }
        }

        /// <summary>
        /// Gets and Set the password.
        /// </summary>
        public string Password
        {
            get
            {
                return _password;
            }
            set
            {
                _password = value;
            }
        }

        /// <summary>
        /// If the value of mode is true, set binary mode for downloads, else, Ascii mode.
        /// </summary>
        public bool BinaryMode
        {
            get
            {
                return BinMode;
            }
            set
            {
                if (BinMode == value) return;

                SendCommand(value ? "TYPE I" : "TYPE A");

                if (_resultCode != 200) throw new FtpException(_result.Substring(4));
            }
        }

        /// <summary>
        /// Login to the remote server.
        /// </summary>
        public void Login()
        {
            if (_loggedin) Close();

            Debug.WriteLine("Opening connection to " + _server, "FtpClient");

            try
            {
                _clientSocket = new Socket(AddressFamily.InterNetwork, SocketType.Stream, ProtocolType.Tcp);
                IPAddress addr = Dns.GetHostEntry(_server).AddressList[0];
                IPEndPoint ep = new IPEndPoint(addr, _port);
                _clientSocket.Connect(ep);
            }
            catch (Exception ex)
            {
                // doubtfull
                if (_clientSocket != null && _clientSocket.Connected) _clientSocket.Close();

                throw new FtpException("Couldn't connect to remote server", ex);
            }

            ReadResponse();

            if (_resultCode != 220)
            {
                Close();
                throw new FtpException(_result.Substring(4));
            }

            SendCommand("USER " + _username);

            if (!(_resultCode == 331 || _resultCode == 230))
            {
                Cleanup();
                throw new FtpException(_result.Substring(4));
            }

            if (_resultCode != 230)
            {
                SendCommand("PASS " + _password);

                if (!(_resultCode == 230 || _resultCode == 202))
                {
                    Cleanup();
                    throw new FtpException(_result.Substring(4));
                }
            }

            _loggedin = true;

            Debug.WriteLine("Connected to " + _server, "FtpClient");

            ChangeDir(_remotePath);
        }

        /// <summary>
        /// Close the FTP connection.
        /// </summary>
        public void Close()
        {
            Debug.WriteLine("Closing connection to " + _server, "FtpClient");

            if (_clientSocket != null)
            {
                SendCommand("QUIT");
            }

            Cleanup();
        }

        /// <summary>
        /// Return a string array containing the remote directory's file list.
        /// </summary>
        /// <returns></returns>
        public string[] GetFileList()
        {
            return GetFileList("*.*");
        }

        /// <summary>
        /// Return a string array containing the remote directory's file list.
        /// </summary>
        /// <param name="mask"></param>
        /// <returns></returns>
        public string[] GetFileList(string mask)
        {
            if (!_loggedin) Login();

            Socket cSocket = CreateDataSocket();

            SendCommand("NLST " + mask);

            if (!(_resultCode == 150 || _resultCode == 125)) throw new FtpException(_result.Substring(4));

            _message = "";

            DateTime timeout = DateTime.Now.AddSeconds(_timeoutSeconds);

            while (timeout > DateTime.Now)
            {
                int bytes = cSocket.Receive(_buffer, _buffer.Length, 0);
                _message += Ascii.GetString(_buffer, 0, bytes);

                if (bytes < _buffer.Length) break;
            }

            string[] msg = _message.Replace("\r", "").Split('\n');

            cSocket.Close();

            if (_message.IndexOf("No such file or directory", StringComparison.Ordinal) != -1)
                msg = new string[] { };

            ReadResponse();

            if (_resultCode != 226)
                msg = new string[] { };
            //	throw new FtpException(result.Substring(4));

            return msg;
        }

        /// <summary>
        /// Return the size of a file.
        /// </summary>
        /// <param name="fileName"></param>
        /// <returns></returns>
        public long GetFileSize(string fileName)
        {
            if (!_loggedin) Login();

            SendCommand("SIZE " + fileName);
            long size;

            if (_resultCode == 213)
                size = long.Parse(_result.Substring(4));

            else
                throw new FtpException(_result.Substring(4));

            return size;
        }

        /// <summary>
        /// Download a file to the Assembly's local directory,
        /// keeping the same file name.
        /// </summary>
        /// <param name="remFileName"></param>
        public void Download(string remFileName)
        {
            Download(remFileName, "", false);
        }

        /// <summary>
        /// Download a remote file to the Assembly's local directory,
        /// keeping the same file name, and set the resume flag.
        /// </summary>
        /// <param name="remFileName"></param>
        /// <param name="resume"></param>
        public void Download(string remFileName, Boolean resume)
        {
            Download(remFileName, "", resume);
        }

        /// <summary>
        /// Download a remote file to a local file name which can include
        /// a path. The local file name will be created or overwritten,
        /// but the path must exist.
        /// </summary>
        /// <param name="remFileName"></param>
        /// <param name="locFileName"></param>
        public void Download(string remFileName, string locFileName)
        {
            Download(remFileName, locFileName, false);
        }

        /// <summary>
        /// Download a remote file to a local file name which can include
        /// a path, and set the resume flag. The local file name will be
        /// created or overwritten, but the path must exist.
        /// </summary>
        /// <param name="remFileName"></param>
        /// <param name="locFileName"></param>
        /// <param name="resume"></param>
        public void Download(string remFileName, string locFileName, Boolean resume)
        {
            if (!_loggedin) Login();

            BinaryMode = true;

            Debug.WriteLine("Downloading file " + remFileName + " from " + _server + "/" + _remotePath, "FtpClient");

            if (locFileName.Equals(""))
            {
                locFileName = remFileName;
            }

            FileStream output = !File.Exists(locFileName) ? File.Create(locFileName) : new FileStream(locFileName, FileMode.Open);

            Socket cSocket = CreateDataSocket();

            if (resume)
            {
                long offset = output.Length;

                if (offset > 0)
                {
                    SendCommand("REST " + offset);
                    if (_resultCode != 350)
                    {
                        //Server dosnt support resuming
                        offset = 0;
                        Debug.WriteLine("Resuming not supported:" + _result.Substring(4), "FtpClient");
                    }
                    else
                    {
                        Debug.WriteLine("Resuming at offset " + offset, "FtpClient");
                        output.Seek(offset, SeekOrigin.Begin);
                    }
                }
            }

            SendCommand("RETR " + remFileName);

            if (_resultCode != 150 && _resultCode != 125)
            {
                throw new FtpException(_result.Substring(4));
            }

            DateTime timeout = DateTime.Now.AddSeconds(_timeoutSeconds);

            while (timeout > DateTime.Now)
            {
                _bytes = cSocket.Receive(_buffer, _buffer.Length, 0);
                output.Write(_buffer, 0, _bytes);

                if (_bytes <= 0)
                {
                    break;
                }
            }

            output.Close();

            if (cSocket.Connected) cSocket.Close();

            ReadResponse();

            if (_resultCode != 226 && _resultCode != 250)
                throw new FtpException(_result.Substring(4));
        }

        /// <summary>
        /// Upload a file.
        /// </summary>
        /// <param name="fileName"></param>
        public void Upload(string fileName)
        {
            Upload(fileName, false);
        }

        /// <summary>
        /// Upload a file and set the resume flag.
        /// </summary>
        /// <param name="fileName"></param>
        /// <param name="resume"></param>
        public void Upload(string fileName, bool resume)
        {
            if (!_loggedin) Login();

            long offset = 0;

            if (resume)
            {
                try
                {
                    BinaryMode = true;

                    offset = GetFileSize(Path.GetFileName(fileName));
                }
                catch (Exception ex)
                {
                    throw (new Exception("File " + fileName + " does not exist.Ex:  " + ex.Message));
                    //offset = 0;
                }
            }

            // open stream to read file
            FileStream input = new FileStream(fileName, FileMode.Open);

            if (resume && input.Length < offset)
            {
                // different file size
                Debug.WriteLine("Overwriting " + fileName, "FtpClient");
                offset = 0;
            }
            else if (resume && input.Length == offset)
            {
                // file done
                input.Close();
                Debug.WriteLine("Skipping completed " + fileName + " - turn resume off to not detect.", "FtpClient");
                return;
            }

            // dont create untill we know that we need it
            Socket cSocket = CreateDataSocket();

            if (offset > 0)
            {
                SendCommand("REST " + offset);
                if (_resultCode != 350)
                {
                    Debug.WriteLine("Resuming not supported", "FtpClient");
                    offset = 0;
                }
            }

            SendCommand("STOR " + Path.GetFileName(fileName));

            if (_resultCode != 125 && _resultCode != 150) throw new FtpException(_result.Substring(4));

            if (offset != 0)
            {
                Debug.WriteLine("Resuming at offset " + offset, "FtpClient");

                input.Seek(offset, SeekOrigin.Begin);
            }

            Debug.WriteLine("Uploading file " + fileName + " to " + _remotePath, "FtpClient");

            while ((_bytes = input.Read(_buffer, 0, _buffer.Length)) > 0)
            {
                cSocket.Send(_buffer, _bytes, 0);
            }

            input.Close();

            if (cSocket.Connected)
            {
                cSocket.Close();
            }

            ReadResponse();

            if (_resultCode != 226 && _resultCode != 250) throw new FtpException(_result.Substring(4));
        }

        /// <summary>
        /// Upload a directory and its file contents
        /// </summary>
        /// <param name="path"></param>
        /// <param name="recurse">Whether to recurse sub directories</param>
        public void UploadDirectory(string path, bool recurse)
        {
            UploadDirectory(path, recurse, "*.*");
        }

        /// <summary>
        /// Upload a directory and its file contents
        /// </summary>
        /// <param name="path"></param>
        /// <param name="recurse">Whether to recurse sub directories</param>
        /// <param name="mask">Only upload files of the given mask - everything is '*.*'</param>
        public void UploadDirectory(string path, bool recurse, string mask)
        {
            string[] dirs = path.Replace("/", @"\").Split('\\');
            string rootDir = dirs[dirs.Length - 1];

            // make the root dir if it doed not exist
            if (GetFileList(rootDir).Length < 1) MakeDir(rootDir);

            ChangeDir(rootDir);

            foreach (string file in Directory.GetFiles(path, mask))
            {
                Upload(file, true);
            }
            if (recurse)
            {
                foreach (string directory in Directory.GetDirectories(path))
                {
                    UploadDirectory(directory, true, mask);
                }
            }

            ChangeDir("..");
        }

        /// <summary>
        /// Delete a file from the remote FTP server.
        /// </summary>
        /// <param name="fileName"></param>
        public void DeleteFile(string fileName)
        {
            if (!_loggedin) Login();

            SendCommand("DELE " + fileName);

            if (_resultCode != 250) throw new FtpException(_result.Substring(4));

            Debug.WriteLine("Deleted file " + fileName, "FtpClient");
        }

        /// <summary>
        /// Rename a file on the remote FTP server.
        /// </summary>
        /// <param name="oldFileName"></param>
        /// <param name="newFileName"></param>
        /// <param name="overwrite">setting to false will throw exception if it exists</param>
        public void RenameFile(string oldFileName, string newFileName, bool overwrite)
        {
            if (!_loggedin) Login();

            SendCommand("RNFR " + oldFileName);

            if (_resultCode != 350) throw new FtpException(_result.Substring(4));

            if (!overwrite && GetFileList(newFileName).Length > 0) throw new FtpException("File already exists");

            SendCommand("RNTO " + newFileName);

            if (_resultCode != 250) throw new FtpException(_result.Substring(4));

            Debug.WriteLine("Renamed file " + oldFileName + " to " + newFileName, "FtpClient");
        }

        /// <summary>
        /// Create a directory on the remote FTP server.
        /// </summary>
        /// <param name="dirName"></param>
        public void MakeDir(string dirName)
        {
            if (!_loggedin) Login();

            SendCommand("MKD " + dirName);

            if (_resultCode != 250 && _resultCode != 257) throw new FtpException(_result.Substring(4));

            Debug.WriteLine("Created directory " + dirName, "FtpClient");
        }

        /// <summary>
        /// Delete a directory on the remote FTP server.
        /// </summary>
        /// <param name="dirName"></param>
        public void RemoveDir(string dirName)
        {
            if (!_loggedin) Login();

            SendCommand("RMD " + dirName);

            if (_resultCode != 250) throw new FtpException(_result.Substring(4));

            Debug.WriteLine("Removed directory " + dirName, "FtpClient");
        }

        /// <summary>
        /// Change the current working directory on the remote FTP server.
        /// </summary>
        /// <param name="dirName"></param>
        public void ChangeDir(string dirName)
        {
            if (dirName == null || dirName.Equals(".") || dirName.Length == 0)
            {
                return;
            }

            if (!_loggedin) Login();

            SendCommand("CWD " + dirName);

            if (_resultCode != 250) throw new FtpException(_result.Substring(4));

            SendCommand("PWD");

            if (_resultCode != 257) throw new FtpException(_result.Substring(4));

            // gonna have to do better than ...
            _remotePath = _message.Split('"')[1];

            Debug.WriteLine("Current directory is " + _remotePath, "FtpClient");
        }

        /// <summary>
        ///
        /// </summary>
        private void ReadResponse()
        {
            _message = "";
            _result = ReadLine();

            if (_result.Length > 3)
                _resultCode = int.Parse(_result.Substring(0, 3));
            else
                _result = null;
        }

        /// <summary>
        ///
        /// </summary>
        /// <returns></returns>
        private string ReadLine()
        {
            while (true)
            {
                _bytes = _clientSocket.Receive(_buffer, _buffer.Length, 0);
                _message += Ascii.GetString(_buffer, 0, _bytes);

                if (_bytes < _buffer.Length)
                {
                    break;
                }
            }

            string[] msg = _message.Split('\n');

            _message = _message.Length > 2 ? msg[msg.Length - 2] : msg[0];

            if (_message.Length > 4 && !_message.Substring(3, 1).Equals(" ")) return ReadLine();

            if (!_verboseDebugging) return _message;
            for (int i = 0; i < msg.Length - 1; i++)
            {
                Debug.Write(msg[i], "FtpClient");
            }

            return _message;
        }

        /// <summary>
        ///
        /// </summary>
        /// <param name="command"></param>
        private void SendCommand(String command)
        {
            if (_verboseDebugging) Debug.WriteLine(command, "FtpClient");

            Byte[] cmdBytes = Encoding.ASCII.GetBytes((command + "\r\n").ToCharArray());
            _clientSocket.Send(cmdBytes, cmdBytes.Length, 0);
            ReadResponse();
        }

        /// <summary>
        /// when doing data transfers, we need to open another socket for it.
        /// </summary>
        /// <returns>Connected socket</returns>
        private Socket CreateDataSocket()
        {
            SendCommand("PASV");

            if (_resultCode != 227) throw new FtpException(_result.Substring(4));

            int index1 = _result.IndexOf('(');
            int index2 = _result.IndexOf(')');

            string ipData = _result.Substring(index1 + 1, index2 - index1 - 1);

            int[] parts = new int[6];

            int len = ipData.Length;
            int partCount = 0;
            string buf = "";

            for (int i = 0; i < len && partCount <= 6; i++)
            {
                char ch = char.Parse(ipData.Substring(i, 1));

                if (char.IsDigit(ch))
                    buf += ch;

                else if (ch != ',')
                    throw new FtpException("Malformed PASV result: " + _result);

                if (ch != ',' && i + 1 != len) continue;
                try
                {
                    parts[partCount++] = int.Parse(buf);
                    buf = "";
                }
                catch (Exception ex)
                {
                    throw new FtpException("Malformed PASV result (not supported?): " + _result, ex);
                }
            }

            string ipAddress = parts[0] + "." + parts[1] + "." + parts[2] + "." + parts[3];

            int port = (parts[4] << 8) + parts[5];

            Socket socket = null;

            try
            {
                socket = new Socket(AddressFamily.InterNetwork, SocketType.Stream, ProtocolType.Tcp);
                IPEndPoint ep = new IPEndPoint(Dns.GetHostEntry(ipAddress).AddressList[0], port);
                socket.Connect(ep);
            }
            catch (Exception ex)
            {
                // doubtfull....
                if (socket != null && socket.Connected) socket.Close();

                throw new FtpException("Can't connect to remote server", ex);
            }

            return socket;
        }

        /// <summary>
        /// Always release those sockets.
        /// </summary>
        private void Cleanup()
        {
            if (_clientSocket != null)
            {
                _clientSocket.Close();
                _clientSocket = null;
            }
            _loggedin = false;
        }

        /// <summary>
        /// Destuctor
        /// </summary>
        ~FtpClient()
        {
            Cleanup();
        }

        //        /**************************************************************************************************************/
        //        #region Async methods (auto generated)

        //        /*
        //                WinInetApi.FtpClient ftp = new WinInetApi.FtpClient();

        //                MethodInfo[] methods = ftp.GetType().GetMethods(BindingFlags.DeclaredOnly|BindingFlags.Instance|BindingFlags.Public);

        //                foreach ( MethodInfo method in methods )
        //                {
        //                    string param = "";
        //                    string values = "";
        //                    foreach ( ParameterInfo i in  method.GetParameters() )
        //                    {
        //                        param += i.ParameterType.Name + " " + i.Name + ",";
        //                        values += i.Name + ",";
        //                    }

        //                    Debug.WriteLine("private delegate " + method.ReturnType.Name + " " + method.Name + "Callback(" + param.TrimEnd(',') + ");");

        //                    Debug.WriteLine("public System.IAsyncResult Begin" + method.Name + "( " + param + " System.AsyncCallback callback )");
        //                    Debug.WriteLine("{");
        //                    Debug.WriteLine("" + method.Name + "Callback ftpCallback = new " + method.Name + "Callback(" + values + " " + method.Name + ");");
        //                    Debug.WriteLine("return ftpCallback.BeginInvoke(callback, null);");
        //                    Debug.WriteLine("}");
        //                    Debug.WriteLine("public void End" + method.Name + "(System.IAsyncResult asyncResult)");
        //                    Debug.WriteLine("{");
        //                    Debug.WriteLine(method.Name + "Callback fc = (" + method.Name + "Callback) ((AsyncResult)asyncResult).AsyncDelegate;");
        //                    Debug.WriteLine("fc.EndInvoke(asyncResult);");
        //                    Debug.WriteLine("}");
        //                    //Debug.WriteLine(method);
        //                }
        //*/

        //        private delegate void LoginCallback();
        //        public System.IAsyncResult BeginLogin(System.AsyncCallback callback)
        //        {
        //            LoginCallback ftpCallback = new LoginCallback(Login);
        //            return ftpCallback.BeginInvoke(callback, null);
        //        }
        //        private delegate void CloseCallback();
        //        public System.IAsyncResult BeginClose(System.AsyncCallback callback)
        //        {
        //            CloseCallback ftpCallback = new CloseCallback(Close);
        //            return ftpCallback.BeginInvoke(callback, null);
        //        }
        //        private delegate String[] GetFileListCallback();
        //        public System.IAsyncResult BeginGetFileList(System.AsyncCallback callback)
        //        {
        //            GetFileListCallback ftpCallback = new GetFileListCallback(GetFileList);
        //            return ftpCallback.BeginInvoke(callback, null);
        //        }
        //        private delegate String[] GetFileListMaskCallback(String mask);
        //        public System.IAsyncResult BeginGetFileList(String mask, System.AsyncCallback callback)
        //        {
        //            GetFileListMaskCallback ftpCallback = new GetFileListMaskCallback(GetFileList);
        //            return ftpCallback.BeginInvoke(mask, callback, null);
        //        }
        //        private delegate Int64 GetFileSizeCallback(String fileName);
        //        public System.IAsyncResult BeginGetFileSize(String fileName, System.AsyncCallback callback)
        //        {
        //            GetFileSizeCallback ftpCallback = new GetFileSizeCallback(GetFileSize);
        //            return ftpCallback.BeginInvoke(fileName, callback, null);
        //        }
        //        private delegate void DownloadCallback(String remFileName);
        //        public System.IAsyncResult BeginDownload(String remFileName, System.AsyncCallback callback)
        //        {
        //            DownloadCallback ftpCallback = new DownloadCallback(Download);
        //            return ftpCallback.BeginInvoke(remFileName, callback, null);
        //        }
        //        private delegate void DownloadFileNameResumeCallback(String remFileName, Boolean resume);
        //        public System.IAsyncResult BeginDownload(String remFileName, Boolean resume, System.AsyncCallback callback)
        //        {
        //            DownloadFileNameResumeCallback ftpCallback = new DownloadFileNameResumeCallback(Download);
        //            return ftpCallback.BeginInvoke(remFileName, resume, callback, null);
        //        }
        //        private delegate void DownloadFileNameFileNameCallback(String remFileName, String locFileName);
        //        public System.IAsyncResult BeginDownload(String remFileName, String locFileName, System.AsyncCallback callback)
        //        {
        //            DownloadFileNameFileNameCallback ftpCallback = new DownloadFileNameFileNameCallback(Download);
        //            return ftpCallback.BeginInvoke(remFileName, locFileName, callback, null);
        //        }
        //        private delegate void DownloadFileNameFileNameResumeCallback(String remFileName, String locFileName, Boolean resume);
        //        public System.IAsyncResult BeginDownload(String remFileName, String locFileName, Boolean resume, System.AsyncCallback callback)
        //        {
        //            DownloadFileNameFileNameResumeCallback ftpCallback = new DownloadFileNameFileNameResumeCallback(Download);
        //            return ftpCallback.BeginInvoke(remFileName, locFileName, resume, callback, null);
        //        }
        //        private delegate void UploadCallback(String fileName);
        //        public System.IAsyncResult BeginUpload(String fileName, System.AsyncCallback callback)
        //        {
        //            UploadCallback ftpCallback = new UploadCallback(Upload);
        //            return ftpCallback.BeginInvoke(fileName, callback, null);
        //        }
        //        private delegate void UploadFileNameResumeCallback(String fileName, Boolean resume);
        //        public System.IAsyncResult BeginUpload(String fileName, Boolean resume, System.AsyncCallback callback)
        //        {
        //            UploadFileNameResumeCallback ftpCallback = new UploadFileNameResumeCallback(Upload);
        //            return ftpCallback.BeginInvoke(fileName, resume, callback, null);
        //        }
        //        private delegate void UploadDirectoryCallback(String path, Boolean recurse);
        //        public System.IAsyncResult BeginUploadDirectory(String path, Boolean recurse, System.AsyncCallback callback)
        //        {
        //            UploadDirectoryCallback ftpCallback = new UploadDirectoryCallback(UploadDirectory);
        //            return ftpCallback.BeginInvoke(path, recurse, callback, null);
        //        }
        //        private delegate void UploadDirectoryPathRecurseMaskCallback(String path, Boolean recurse, String mask);
        //        public System.IAsyncResult BeginUploadDirectory(String path, Boolean recurse, String mask, System.AsyncCallback callback)
        //        {
        //            UploadDirectoryPathRecurseMaskCallback ftpCallback = new UploadDirectoryPathRecurseMaskCallback(UploadDirectory);
        //            return ftpCallback.BeginInvoke(path, recurse, mask, callback, null);
        //        }
        //        private delegate void DeleteFileCallback(String fileName);
        //        public System.IAsyncResult BeginDeleteFile(String fileName, System.AsyncCallback callback)
        //        {
        //            DeleteFileCallback ftpCallback = new DeleteFileCallback(DeleteFile);
        //            return ftpCallback.BeginInvoke(fileName, callback, null);
        //        }
        //        private delegate void RenameFileCallback(String oldFileName, String newFileName, Boolean overwrite);
        //        public System.IAsyncResult BeginRenameFile(String oldFileName, String newFileName, Boolean overwrite, System.AsyncCallback callback)
        //        {
        //            RenameFileCallback ftpCallback = new RenameFileCallback(RenameFile);
        //            return ftpCallback.BeginInvoke(oldFileName, newFileName, overwrite, callback, null);
        //        }
        //        private delegate void MakeDirCallback(String dirName);
        //        public System.IAsyncResult BeginMakeDir(String dirName, System.AsyncCallback callback)
        //        {
        //            MakeDirCallback ftpCallback = new MakeDirCallback(MakeDir);
        //            return ftpCallback.BeginInvoke(dirName, callback, null);
        //        }
        //        private delegate void RemoveDirCallback(String dirName);
        //        public System.IAsyncResult BeginRemoveDir(String dirName, System.AsyncCallback callback)
        //        {
        //            RemoveDirCallback ftpCallback = new RemoveDirCallback(RemoveDir);
        //            return ftpCallback.BeginInvoke(dirName, callback, null);
        //        }
        //        private delegate void ChangeDirCallback(String dirName);
        //        public System.IAsyncResult BeginChangeDir(String dirName, System.AsyncCallback callback)
        //        {
        //            ChangeDirCallback ftpCallback = new ChangeDirCallback(ChangeDir);
        //            return ftpCallback.BeginInvoke(dirName, callback, null);
        //        }

        //        #endregion
    }
}