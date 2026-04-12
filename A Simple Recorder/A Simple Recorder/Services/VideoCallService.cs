using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Sockets;
using System.Threading;
using System.Threading.Tasks;
using Windows.Graphics.Imaging;

namespace A_Simple_Recorder.Services
{
    public enum MessageType : byte
    {
        VideoFrame = 1,
        AudioChunk = 2,
        ClientInfo = 3,
        VoiceActivity = 4,
        Disconnect = 5
    }

    public class ParticipantInfo
    {
        public string Id { get; set; } = string.Empty;
        public string Name { get; set; } = string.Empty;
        public TcpClient? Client { get; set; }
        public DateTime LastFrameTime { get; set; }
        public bool IsSpeaking { get; set; }
        public float AudioLevel { get; set; }
    }

    public class VideoCallService
    {
        private const int PORT = 9876;
        private const int MAX_FRAME_SIZE = 1024 * 1024; // 1MB max frame size

        private TcpListener? _listener;
        private TcpClient? _clientConnection;
        private bool _isRunning;
        private bool _isHost;
        private string _myId;
        private string _myName;
        private CancellationTokenSource? _cancellationTokenSource;

        private readonly Dictionary<string, ParticipantInfo> _participants = new();
        private readonly object _participantsLock = new object();

        public event Action<string, byte[]>? OnVideoFrameReceived;
        public event Action<string, byte[]>? OnAudioChunkReceived;
        public event Action<ParticipantInfo>? OnParticipantJoined;
        public event Action<string>? OnParticipantLeft;
        public event Action<string, bool, float>? OnVoiceActivity;

        public VideoCallService()
        {
            _myId = Guid.NewGuid().ToString();
            _myName = Environment.MachineName;
        }

        public bool IsHost => _isHost;
        public bool IsConnected => _isRunning;
        public IEnumerable<ParticipantInfo> Participants
        {
            get
            {
                lock (_participantsLock)
                {
                    return _participants.Values.ToList();
                }
            }
        }

        public async Task<bool> StartHostingAsync()
        {
            if (_isRunning) return false;

            try
            {
                _isHost = true;
                _isRunning = true;
                _cancellationTokenSource = new CancellationTokenSource();

                _listener = new TcpListener(IPAddress.Any, PORT);
                _listener.Start();

                System.Diagnostics.Debug.WriteLine($"Hosting on port {PORT}");

                // Start accepting clients
                _ = Task.Run(() => AcceptClientsAsync(_cancellationTokenSource.Token));

                return true;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error starting host: {ex.Message}");
                _isRunning = false;
                _isHost = false;
                return false;
            }
        }

        public async Task<bool> ConnectToHostAsync(string ipAddress)
        {
            if (_isRunning) return false;

            try
            {
                _isHost = false;
                _isRunning = true;
                _cancellationTokenSource = new CancellationTokenSource();

                _clientConnection = new TcpClient();
                await _clientConnection.ConnectAsync(ipAddress, PORT);

                System.Diagnostics.Debug.WriteLine($"Connected to {ipAddress}:{PORT}");

                // Send client info
                await SendClientInfoAsync(_clientConnection);

                // Start receiving data
                _ = Task.Run(() => ReceiveDataAsync(_clientConnection, _cancellationTokenSource.Token));

                return true;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error connecting to host: {ex.Message}");
                _isRunning = false;
                _clientConnection?.Close();
                return false;
            }
        }

        public async Task DisconnectAsync()
        {
            if (!_isRunning) return;

            _isRunning = false;
            _cancellationTokenSource?.Cancel();

            if (_isHost)
            {
                // Notify all clients
                lock (_participantsLock)
                {
                    foreach (var participant in _participants.Values)
                    {
                        participant.Client?.Close();
                    }
                    _participants.Clear();
                }

                _listener?.Stop();
                _listener = null;
            }
            else
            {
                // Send disconnect message
                if (_clientConnection?.Connected == true)
                {
                    await SendDisconnectAsync(_clientConnection);
                }
                _clientConnection?.Close();
                _clientConnection = null;
            }

            _cancellationTokenSource?.Dispose();
            _cancellationTokenSource = null;

            System.Diagnostics.Debug.WriteLine("Disconnected from video call");
        }

        public async Task SendVideoFrameAsync(byte[] frameData)
        {
            if (!_isRunning || frameData == null || frameData.Length == 0)
                return;

            try
            {
                if (_isHost)
                {
                    // Get copy of clients outside lock
                    List<TcpClient> clients;
                    lock (_participantsLock)
                    {
                        clients = _participants.Values
                            .Where(p => p.Client?.Connected == true)
                            .Select(p => p.Client!)
                            .ToList();
                    }

                    // Broadcast to all clients
                    foreach (var client in clients)
                    {
                        await SendMessageAsync(client, MessageType.VideoFrame, frameData);
                    }
                }
                else if (_clientConnection?.Connected == true)
                {
                    // Send to host
                    await SendMessageAsync(_clientConnection, MessageType.VideoFrame, frameData);
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error sending video frame: {ex.Message}");
            }
        }

        public async Task SendAudioChunkAsync(byte[] audioData)
        {
            if (!_isRunning || audioData == null || audioData.Length == 0)
                return;

            try
            {
                if (_isHost)
                {
                    // Get copy of clients outside lock
                    List<TcpClient> clients;
                    lock (_participantsLock)
                    {
                        clients = _participants.Values
                            .Where(p => p.Client?.Connected == true)
                            .Select(p => p.Client!)
                            .ToList();
                    }

                    // Broadcast to all clients
                    foreach (var client in clients)
                    {
                        await SendMessageAsync(client, MessageType.AudioChunk, audioData);
                    }
                }
                else if (_clientConnection?.Connected == true)
                {
                    // Send to host
                    await SendMessageAsync(_clientConnection, MessageType.AudioChunk, audioData);
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error sending audio chunk: {ex.Message}");
            }
        }

        public async Task SendVoiceActivityAsync(bool isSpeaking, float audioLevel)
        {
            if (!_isRunning) return;

            try
            {
                var data = new byte[5];
                data[0] = isSpeaking ? (byte)1 : (byte)0;
                BitConverter.GetBytes(audioLevel).CopyTo(data, 1);

                if (_isHost)
                {
                    // Get copy of clients outside lock
                    List<TcpClient> clients;
                    lock (_participantsLock)
                    {
                        clients = _participants.Values
                            .Where(p => p.Client?.Connected == true)
                            .Select(p => p.Client!)
                            .ToList();
                    }

                    foreach (var client in clients)
                    {
                        await SendMessageAsync(client, MessageType.VoiceActivity, data);
                    }
                }
                else if (_clientConnection?.Connected == true)
                {
                    await SendMessageAsync(_clientConnection, MessageType.VoiceActivity, data);
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error sending voice activity: {ex.Message}");
            }
        }

        private async Task AcceptClientsAsync(CancellationToken cancellationToken)
        {
            while (_isRunning && !cancellationToken.IsCancellationRequested)
            {
                try
                {
                    var client = await _listener!.AcceptTcpClientAsync();
                    
                    lock (_participantsLock)
                    {
                        if (_participants.Count >= 4) // Max 4 clients (+ 1 host = 5 total)
                        {
                            client.Close();
                            continue;
                        }
                    }

                    System.Diagnostics.Debug.WriteLine($"Client connected: {client.Client.RemoteEndPoint}");

                    // Start receiving from this client
                    _ = Task.Run(() => ReceiveDataAsync(client, cancellationToken), cancellationToken);
                }
                catch (Exception ex)
                {
                    if (_isRunning)
                    {
                        System.Diagnostics.Debug.WriteLine($"Error accepting client: {ex.Message}");
                    }
                }
            }
        }

        private async Task ReceiveDataAsync(TcpClient client, CancellationToken cancellationToken)
        {
            string? clientId = null;

            try
            {
                var stream = client.GetStream();
                var buffer = new byte[8192];

                while (_isRunning && !cancellationToken.IsCancellationRequested && client.Connected)
                {
                    // Read message header: [Type:1][Size:4][SenderId:36]
                    var headerBuffer = new byte[41];
                    var headerRead = await ReadExactAsync(stream, headerBuffer, 0, 41, cancellationToken);
                    if (headerRead != 41) break;

                    var messageType = (MessageType)headerBuffer[0];
                    var dataSize = BitConverter.ToInt32(headerBuffer, 1);
                    clientId = System.Text.Encoding.UTF8.GetString(headerBuffer, 5, 36);

                    if (dataSize > MAX_FRAME_SIZE || dataSize < 0)
                    {
                        System.Diagnostics.Debug.WriteLine($"Invalid data size: {dataSize}");
                        break;
                    }

                    // Read message data
                    var data = new byte[dataSize];
                    var dataRead = await ReadExactAsync(stream, data, 0, dataSize, cancellationToken);
                    if (dataRead != dataSize) break;

                    // Process message
                    await ProcessMessageAsync(clientId, client, messageType, data);
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error receiving data: {ex.Message}");
            }
            finally
            {
                if (clientId != null)
                {
                    lock (_participantsLock)
                    {
                        if (_participants.Remove(clientId))
                        {
                            OnParticipantLeft?.Invoke(clientId);
                            System.Diagnostics.Debug.WriteLine($"Participant left: {clientId}");
                        }
                    }
                }
                client.Close();
            }
        }

        private async Task<int> ReadExactAsync(NetworkStream stream, byte[] buffer, int offset, int count, CancellationToken cancellationToken)
        {
            int totalRead = 0;
            while (totalRead < count && !cancellationToken.IsCancellationRequested)
            {
                var read = await stream.ReadAsync(buffer, offset + totalRead, count - totalRead, cancellationToken);
                if (read == 0) break;
                totalRead += read;
            }
            return totalRead;
        }

        private async Task ProcessMessageAsync(string senderId, TcpClient senderClient, MessageType messageType, byte[] data)
        {
            switch (messageType)
            {
                case MessageType.ClientInfo:
                    var name = System.Text.Encoding.UTF8.GetString(data);
                    var participant = new ParticipantInfo
                    {
                        Id = senderId,
                        Name = name,
                        Client = senderClient,
                        LastFrameTime = DateTime.Now
                    };

                    lock (_participantsLock)
                    {
                        _participants[senderId] = participant;
                    }

                    OnParticipantJoined?.Invoke(participant);
                    System.Diagnostics.Debug.WriteLine($"Participant joined: {name} ({senderId})");
                    break;

                case MessageType.VideoFrame:
                    OnVideoFrameReceived?.Invoke(senderId, data);
                    
                    lock (_participantsLock)
                    {
                        if (_participants.TryGetValue(senderId, out var p))
                        {
                            p.LastFrameTime = DateTime.Now;
                        }
                    }

                    // If host, broadcast to other clients
                    if (_isHost)
                    {
                        await BroadcastToOthersAsync(senderId, MessageType.VideoFrame, data);
                    }
                    break;

                case MessageType.AudioChunk:
                    OnAudioChunkReceived?.Invoke(senderId, data);

                    // If host, broadcast to other clients
                    if (_isHost)
                    {
                        await BroadcastToOthersAsync(senderId, MessageType.AudioChunk, data);
                    }
                    break;

                case MessageType.VoiceActivity:
                    var isSpeaking = data[0] == 1;
                    var audioLevel = BitConverter.ToSingle(data, 1);
                    
                    lock (_participantsLock)
                    {
                        if (_participants.TryGetValue(senderId, out var p))
                        {
                            p.IsSpeaking = isSpeaking;
                            p.AudioLevel = audioLevel;
                        }
                    }

                    OnVoiceActivity?.Invoke(senderId, isSpeaking, audioLevel);

                    // If host, broadcast to other clients
                    if (_isHost)
                    {
                        await BroadcastToOthersAsync(senderId, MessageType.VoiceActivity, data);
                    }
                    break;

                case MessageType.Disconnect:
                    // Client wants to disconnect
                    break;
            }
        }

        private async Task BroadcastToOthersAsync(string excludeId, MessageType messageType, byte[] data)
        {
            // Get copy of participants outside lock
            List<ParticipantInfo> participants;
            lock (_participantsLock)
            {
                participants = _participants.Values.ToList();
            }

            foreach (var participant in participants)
            {
                if (participant.Id != excludeId && participant.Client?.Connected == true)
                {
                    await SendMessageAsync(participant.Client, messageType, data, excludeId);
                }
            }
        }

        private async Task SendMessageAsync(TcpClient client, MessageType messageType, byte[] data, string? senderId = null)
        {
            try
            {
                var stream = client.GetStream();
                
                // Message format: [Type:1][Size:4][SenderId:36][Data:n]
                var header = new byte[41];
                header[0] = (byte)messageType;
                BitConverter.GetBytes(data.Length).CopyTo(header, 1);
                
                var idBytes = System.Text.Encoding.UTF8.GetBytes((senderId ?? _myId).PadRight(36));
                Array.Copy(idBytes, 0, header, 5, Math.Min(36, idBytes.Length));

                await stream.WriteAsync(header, 0, header.Length);
                await stream.WriteAsync(data, 0, data.Length);
                await stream.FlushAsync();
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error sending message: {ex.Message}");
            }
        }

        private async Task SendClientInfoAsync(TcpClient client)
        {
            var nameBytes = System.Text.Encoding.UTF8.GetBytes(_myName);
            await SendMessageAsync(client, MessageType.ClientInfo, nameBytes);
        }

        private async Task SendDisconnectAsync(TcpClient client)
        {
            await SendMessageAsync(client, MessageType.Disconnect, Array.Empty<byte>());
        }

        public string GetLocalIPAddress()
        {
            try
            {
                var host = Dns.GetHostEntry(Dns.GetHostName());
                foreach (var ip in host.AddressList)
                {
                    if (ip.AddressFamily == AddressFamily.InterNetwork)
                    {
                        return ip.ToString();
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error getting local IP: {ex.Message}");
            }
            return "127.0.0.1";
        }
    }
}
