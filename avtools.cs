using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Net.Sockets;
using System.IO;
using System.Security.Cryptography;

namespace Enkadia
{
    namespace Projectors
    {
        namespace PJLink
        {
            /// <summary>
            /// Enkadia AV Component for PJLink enabled projectors.
            /// <para>This version does NOT support passwords and user authentication.</para>
            /// <para>Disable administrative password using manufacturer web browser</para>
            /// <para>It is recommended that port 4352 indicated in the PJLink specifications is used to access the projector or display.</para>
            /// </summary>
            public class Projector
            {
                #region Declarations

                string _cmd;
                string _response;
                int _inputSelected;
                string deviceError = "The projector is not connected. \r\n 1. Check projector IP address in the projector network menu. \r\n 2. Check IP address in programming. \r\n 3. Check network connections. \r\n 4. Check network switch";
                //string inputExceedsError = "The input is required to be a number between 1 and 9";

                #endregion

                #region authentication declarations

                /// <summary>
                /// Flag is true, if the projector requires authentication
                /// </summary>
                private bool _useAuth = false;
                /// <summary>
                /// Password supplied by user if authentication is enabled
                /// </summary>
                private string _passwd = "";
                /// <summary>
                /// Random key returned by projector for MD5 sum generation 
                /// </summary>
                private string _pjKey = "";

                /// <summary>
                /// The connection client
                /// </summary>
                TcpClient client = null;
                /// <summary>
                /// The Network stream the _client provides
                /// </summary>
                NetworkStream stream = null;

                #endregion
                #region Properties

                bool IsConnected { get; set; }
                /// <summary>
                /// IPv4 Address assigned by administrator
                /// </summary>
                public string IPAddress { get; set; }
                /// <summary>
                /// According to the PJLink standard the default port is 4352. It is recommended that this port should not be changed
                /// </summary>
                public int Port { get; set; }
                /// <summary>
                /// Shows the current power condition
                /// </summary>
                public int PowerStatus { get; private set; }
                /// <summary>
                /// The current mute state. Off = shutter open, On = shutter closed
                /// </summary>
                public string MuteStatus { get; private set; }
                /// <summary>
                /// Current Error message
                /// </summary>
                public string ErrorMessage { get; private set; }
                string PJKey { get; set; }
                /// <summary>
                /// Used for general projector response messages
                /// </summary>
                public string ProjectorResponse { get; private set; }
                string _authenticationResponse { get; set; }

                #endregion

                #region Enumerations

                public enum Inputs
                {
                    RGB1 = 11,
                    RGB2 = 12,
                    Video = 21,
                    SVideo = 22,
                    DVI = 31,
                    HDMI = 32
                }

                public enum InputType
                {
                    RGB = 1,
                    Video = 2,
                    Digital = 3,
                    Storage = 4,
                    Network = 5
                }

                public enum ProjectorPowerStatus
                {
                    PowerOff = 0,
                    PowerOn = 1,
                    Cooling = 2,
                    WarmUp = 3,
                    Unavailable = 4,
                    ProjectorFailure = 5
                }

                #endregion

                #region SendCommand - Contains connection to projector

                private string SendCommand(string Command)
                {
                    _cmd = Command;

                    if (initConnection())
                    {
                        try
                        {

                            if (_useAuth && _pjKey != "")
                                _cmd = getMD5Hash(_pjKey + _passwd) + _cmd;

                            byte[] sendCommand = Encoding.ASCII.GetBytes(_cmd);
                            stream.Write(sendCommand, 0, sendCommand.Length);

                            byte[] recvBytes = new byte[client.ReceiveBufferSize];
                            int bytesRcvd = stream.Read(recvBytes, 0, (int)client.ReceiveBufferSize);
                            _response = Encoding.ASCII.GetString(recvBytes, 0, bytesRcvd);
                            ProjectorResponse = _response.Trim("\r\0".ToCharArray());
                            return ProjectorResponse;
                        }
                        catch (Exception e)
                        {

                            ErrorMessage = e.Message;
                        }
                        finally
                        {
                            closeConnection();
                        }

                        return ErrorMessage;
                    }

                    return deviceError;
                }

                #endregion

                #region Get Projector Response

                private void GetResponse(string Command)
                {
                    _cmd = Command;

                    client = new TcpClient();
                    

                    client.ConnectAsync(IPAddress, Port).Wait(1000);
                    if(client.Connected == true)
                    {
                        { IsConnected = true; }
                    }


                    try
                    {
                        stream = client.GetStream();
                        byte[] bytes = new byte[client.ReceiveBufferSize = 8088];
                        stream.Read(bytes, 0, (int)client.ReceiveBufferSize);

                        _response = (Encoding.ASCII.GetString(bytes));
                        _response = _response.Trim("\r\0".ToCharArray());

                        // determine if projector is set to use security 0 = no security, 1 = use security
                        if (_response == "PJLINK 0")
                        {
                            _useAuth = false;
                            byte[] sendCommand = Encoding.ASCII.GetBytes(_cmd);
                            stream.Write(sendCommand, 0, sendCommand.Length);
                        }
                        else
                        {
                            _useAuth = true;
                            _pjKey = _response.Replace("PJLINK 1 ", "");

                            if (_useAuth && _pjKey != "")
                                _cmd = getMD5Hash(_pjKey + _passwd) + _cmd;

                            byte[] sendCommand = Encoding.ASCII.GetBytes(_cmd);
                            stream.Write(sendCommand, 0, sendCommand.Length);
                        }

                        stream.Read(bytes, 0, (int)client.ReceiveBufferSize);

                        _response = (Encoding.ASCII.GetString(bytes));
                        _response = _response.Trim("\r\0".ToCharArray());

                        //wrap it up
                        closeConnection();

                    }
                    catch (Exception)
                    {

                    }

                    ProjectorResponse = _response;
                }

                #endregion

                #region Power Settings

                /// <summary>
                /// Start the projector
                /// <example>
                /// <code>
                /// using Enkadia.Panasonic.PJLink
                /// Projector projector = new Projector()
                /// ...
                /// void btnPowerOn_Click{object sender, EventArgs e)
                /// {
                ///     projector.PowerOn(); 
                /// }
                /// </code>
                /// </example>
                /// </summary>
                public void PowerOn()
                {
                    _cmd = "%1POWR 1\r";
                    SendCommand(_cmd);
                }
                /// <summary>
                /// Shutdown the projector
                /// <example>
                /// <code>
                /// using Enkadia.Panasonic.Projectors.PJLink
                /// Projector projector = new Projector()
                /// ...
                /// void btnPowerOn_Click{object sender, EventArgs e)
                /// {
                ///     projector.PowerOff(); 
                /// }
                /// </code>
                /// </example>
                /// </summary>
                public void PowerOff()
                {
                    _cmd = "%1POWR 0\r";
                    SendCommand(_cmd);
                }
                /// <summary>
                /// Check the projector power status
                /// </summary>
                /// <returns>One of six possible states: Power Off, Power On, Cooling Down, Warming Up, Unavailable Time or Projector Failure</returns>
                public int GetPowerStatus()
                {

                    _cmd = "%1POWR ?\r";
                    GetResponse(_cmd);
                    //string _powerStatus = ProjectorResponse;
                    switch (ProjectorResponse)
                    {
                        case "%1POWR=0":
                            PowerStatus = (int)ProjectorPowerStatus.PowerOff;
                            break;
                        case "%1POWR=1":
                            PowerStatus = (int)ProjectorPowerStatus.PowerOn;
                            break;
                        case "%1POWR=2":
                            PowerStatus = (int)ProjectorPowerStatus.Cooling;
                            break;
                        case "%1POWR=3":
                            PowerStatus = (int)ProjectorPowerStatus.WarmUp;
                            break;
                        case "%1POWR=ERR3":
                            PowerStatus = (int)ProjectorPowerStatus.Unavailable;
                            break;
                        case "%1POWR=ERR4":
                            PowerStatus = (int)ProjectorPowerStatus.ProjectorFailure;
                            break;
                        default:
                            break;
                    }

                    return PowerStatus;
                }

                #endregion

                #region InputSelection & Query
                /// <summary>
                /// Check which input is currently selected
                /// </summary>
                /// <returns>Input number based on PJLink standard. e.g. 31 could be Digital In 1, HDMI, DVI or DisplayPort.
                /// Check the manufacturer information for your model of projector or display. </returns>
                public string GetInputSelected()
                {
                    _cmd = "%1INPT ?\r";
                    GetResponse(_cmd);
                    int index = ProjectorResponse.LastIndexOf("=");
                    ProjectorResponse = ProjectorResponse.Substring(index + 1);
                    return ProjectorResponse;
                }


                public void SetInput(int InputType, int Input)
                {
                    StringBuilder sb = new StringBuilder();

                    switch (InputType)
                    {
                        case (int)Projector.InputType.RGB:
                            _inputSelected = (int)Projector.InputType.RGB + Input;
                            break;
                        case (int)Projector.InputType.Video:
                            _inputSelected = (int)Projector.InputType.Video + Input;
                            break;
                        case (int)Projector.InputType.Digital:
                            _inputSelected = (int)Projector.InputType.Digital + Input;
                            break;
                        case (int)Projector.InputType.Storage:
                            _inputSelected = (int)Projector.InputType.Storage + Input;
                            break;
                        case (int)Projector.InputType.Network:
                            _inputSelected = (int)Projector.InputType.Network + Input;
                            break;
                    }

                    sb.Append(Convert.ToString(_inputSelected));
                    sb.Append("\r");

                    _cmd = sb.ToString();
                    SendCommand(_cmd);
                }

                public void InputRGB(int Input)
                {
                    StringBuilder sb = new StringBuilder();
                    sb.Append("%1INPT ");
                    sb.Append((int)InputType.RGB);
                    sb.Append(Input);
                    sb.Append("\r");
                    _cmd = sb.ToString();
                    SendCommand(_cmd);
                }

                public void InputVideo(int Input)
                {

                    StringBuilder sb = new StringBuilder();
                    sb.Append("%1INPT ");
                    sb.Append((int)InputType.Video);
                    sb.Append(Input);
                    sb.Append("\r");
                    _cmd = sb.ToString();
                    SendCommand(_cmd);
                }

                public void InputDigital(int Input)
                {
                    StringBuilder sb = new StringBuilder();
                    sb.Append("%1INPT ");
                    sb.Append((int)InputType.Digital);
                    sb.Append(Input);
                    sb.Append("\r");
                    _cmd = sb.ToString();
                    SendCommand(_cmd);
                }

                public void InputStorage(int Input)
                {
                    StringBuilder sb = new StringBuilder();
                    sb.Append("%1INPT ");
                    sb.Append((int)InputType.Storage);
                    sb.Append(Input);
                    sb.Append("\r");
                    _cmd = sb.ToString();
                    SendCommand(_cmd);
                }

                public void InputNetwork(int Input)
                {
                    StringBuilder sb = new StringBuilder();
                    sb.Append("%1INPT ");
                    sb.Append((int)InputType.Network);
                    sb.Append(Input);
                    sb.Append("\r");
                    _cmd = sb.ToString();
                    SendCommand(_cmd);
                }



                #endregion

                #region AV Mute

                public void AVShutterClose()
                {
                    _cmd = "%1AVMT 31\r";
                    SendCommand(_cmd);
                }

                public void AVShutterOpen()
                {
                    _cmd = "%1AVMT 30\r";
                    SendCommand(_cmd);
                }

                public void AudioMuteOn()
                {
                    _cmd = "%1AVMT 21\r";
                    SendCommand(_cmd);
                }

                public void AudioMuteOff()
                {
                    _cmd = "%1AVMT 20\r";
                    SendCommand(_cmd);
                }

                public void VideoMuteOn()
                {
                    _cmd = "%1AVMT 21\r";
                    SendCommand(_cmd);
                }

                /// <summary>
                /// Check mute status
                /// </summary>
                /// <returns>Returns eight possible states: Video Mute Off/On, Audio Mute Off/On, A/V Mute Off/On, Unavailable/In Standby, Projector Failure</returns>
                public string AVMuteStatus()
                {
                    _cmd = "%1AVMT ?\r";
                    GetResponse(_cmd);
                    MuteStatus = ProjectorResponse;


                    switch (MuteStatus)
                    {
                        case "%1AVMT=10":
                            MuteStatus = "Video Mute Off";
                            break;
                        case "%1AVMT=11":
                            MuteStatus = "Video Mute On";
                            break;
                        case "%1AVMT=20":
                            MuteStatus = "Audio Mute Off";
                            break;
                        case "%1AVMT=21":
                            MuteStatus = "Audio Mute On";
                            break;
                        case "%1AVMT=30":
                            MuteStatus = "Shutter Open";
                            break;
                        case "%1AVMT=31":
                            MuteStatus = "Shutter Closed";
                            break;
                        case "%1AVMT=ERR3":
                            MuteStatus = "Unavailable/In Standby";
                            break;
                        case "%1AVMT=ERR4":
                            MuteStatus = "Projector Failure";
                            break;
                        default:
                            break;
                    }

                    return MuteStatus;
                }


                #endregion

                #region Error Query

                public string GetErrorMessage()
                {
                    _cmd = "%1ERST ?\r";
                    GetResponse(_cmd);
                    ErrorMessage = ProjectorResponse;
                    return ErrorMessage;
                }

                #endregion

                #region Lamp Hours
                /// <summary>
                /// Check the lamp hours and lamp status
                /// </summary>
                /// <returns>1) First five digits indicate cumulative lamp hours. 2) Hours are followed by a space, then a 0 indicating the lamp is off or a 1 indicating it is on.
                /// </returns>
                public string GetLampInformation()
                {
                    _cmd = "%1LAMP ?\r";
                    GetResponse(_cmd);
                    int index = ProjectorResponse.LastIndexOf("=");
                    ProjectorResponse = ProjectorResponse.Substring(index + 1);
                    return ProjectorResponse;
                }

                #endregion

                #region Other Information

                /// <summary>
                /// The projector name must be set locally through the projector interface.
                /// </summary>
                /// <returns>Projector name</returns>
                /// 
                public string GetProjectorName()
                {
                    _cmd = "%1NAME ?\r";
                    GetResponse(_cmd);
                    int index = ProjectorResponse.LastIndexOf("=");
                    ProjectorResponse = ProjectorResponse.Substring(index + 1);
                    return ProjectorResponse;
                }

                /// <summary>
                /// Query typically returns the manufacturer name, but can be used for general information.
                /// </summary>
                /// <returns>Manufacturer name or information found in INF1 location.</returns>
                public string GetManufacturer()
                {
                    _cmd = "%1INF1 ?\r";
                    GetResponse(_cmd);
                    int index = ProjectorResponse.LastIndexOf("=");
                    ProjectorResponse = ProjectorResponse.Substring(index + 1);
                    return ProjectorResponse;
                }
                /// <summary>
                /// Query typically returns the model name and number, but can be used for general information.
                /// </summary>
                /// <returns>Model or information found in INF2 location.</returns>
                public string GetModel()
                {
                    _cmd = "%1INF2 ?\r";
                    GetResponse(_cmd);
                    int index = ProjectorResponse.LastIndexOf("=");
                    ProjectorResponse = ProjectorResponse.Substring(index + 1);
                    return ProjectorResponse;
                }
                #endregion

                #region Authentication Procedure


                private bool initConnection()
                {
                    try
                    {
                        if (client == null || !client.Connected)
                        {
                            client = new TcpClient();
                            client.ConnectAsync(IPAddress, Port).Wait(1000);
                            stream = client.GetStream();
                            byte[] recvBytes = new byte[client.ReceiveBufferSize];
                            int bytesRcvd = stream.Read(recvBytes, 0, (int)client.ReceiveBufferSize);
                            string retVal = Encoding.ASCII.GetString(recvBytes, 0, bytesRcvd);
                            retVal = retVal.Trim();

                            if (retVal.IndexOf("PJLINK 0") >= 0)
                            {
                                _useAuth = false;  //pw provided but projector doesn't need it.
                                return true;
                            }
                            else if (retVal.IndexOf("PJLINK 1 ") >= 0)
                            {
                                _useAuth = true;
                                _pjKey = retVal.Replace("PJLINK 1 ", "");
                                return true;
                            }
                        }
                        return false;
                    }
                    catch (Exception)
                    {
                        return false;
                    }

                }

                //gets the MD5HASH of the concatenated input string(rand_no + password)
                //The random no is from the projector and also the password to the projector
                private string getMD5Hash(string input)
                {
                    MD5CryptoServiceProvider x = new MD5CryptoServiceProvider();
                    byte[] bs = Encoding.ASCII.GetBytes(input);
                    byte[] hash = x.ComputeHash(bs);

                    string toRet = "";
                    foreach (byte b in hash)
                    {
                        toRet += b.ToString("x2");
                    }
                    return toRet;
                }

                private void closeConnection()
                {
                    Thread.Sleep(1500);

                    //wrap it up
                    stream.Flush();
                    stream.Dispose();
                    client.Dispose();
                }

                #endregion
            }
        }

    }

}
