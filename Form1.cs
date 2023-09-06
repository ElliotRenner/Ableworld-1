using System.Diagnostics;
using System.Management;
using System;
using System.Collections.ObjectModel;
using System.Security;
using System.Management.Automation;
using System.Management.Automation.Runspaces;
using static System.Net.Mime.MediaTypeNames;
using static System.Runtime.InteropServices.JavaScript.JSType;
using System.Timers;
using System.Text;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Windows.Forms;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;
using System.Formats.Asn1;


namespace HO_Menu
{

    public partial class Form1 : Form
    {

        private System.Timers.Timer refreshTimer = new System.Timers.Timer();
        //string remoteUsername = ""; // To use in the event that the auth logon breaks
        //string remotePassword = ""; // To use in the event that the auth logon breaks

        public Form1()
        {
            InitializeComponent();
            InitializeTimer();
            RetrieveRemotePCInformation();
        }

        public string remoteUsername { get; set; }
        public string remotePassword { get; set; }

        private void LoginButton_click(object sender, EventArgs e)
        {
            // Retrieve the text from the two TextBox controls
            remoteUsername = LoginUser.Text;
            remotePassword = LoginPassword.Text;
            tabControl1.SelectedTab = Information;
        }

        private void PCName_GotFocus(object sender, EventArgs e)
        {
            // When the text box is selected, stop the timer
            refreshTimer.Stop();
        }

        private void PCName_LostFocus(object sender, EventArgs e)
        {
            // When the text box loses focus, restart the timer
            refreshTimer.Start();
        }

        private void InitializeTimer()
        {
            CPULabel.BackColor = Color.Transparent;
            refreshTimer.Interval = 1000;
            refreshTimer.Elapsed += RefreshTimer_Elapsed;
            refreshTimer.Start();
        }

        private void RefreshTimer_Elapsed(object sender, ElapsedEventArgs e)
        {
            SystemUsage();
        }

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            refreshTimer.Stop();
            refreshTimer.Dispose();
        }
        private string GetWindowsCodename(string version)
        {
            Dictionary<int, string> codenames = new Dictionary<int, string>
    {
        { 23536, "Insider" }, // Windows 11
        { 22631, "23H2" }, // Windows 11
        { 22621, "22H2" }, // Windows 11        
        { 22000, "21H2" }, // Windows 11
        { 19045, "22H2" }, // Windows 10
        { 19044, "21H2" }, // Windows 10
        { 19043, "21H1" }, // Windows 10
        { 19042, "20H2" }, // Windows 10
        // Add more entries as needed
    };

            if (version.Length >= 5 && int.TryParse(version.Substring(version.Length - 5), out int buildNumber))
            {
                if (codenames.ContainsKey(buildNumber))
                {
                    return codenames[buildNumber];
                }
            }

            return "Unknown Build";
        }

        private void button2_Click(object sender, EventArgs e)
        {
            string remoteComputerName = PCName.Text; // Replace with the remote PC's name
            string message = MessageBox.Text;
            OutputBox.Text = "";
            try
            {
                Process.Start("msg", $"/server:{remoteComputerName} * \"{message}\"");
                OutputBox.Text = $"Message sent to {remoteComputerName}";
            }
            catch (Exception ex)
            {
                AddColoredText($" - An error occurred", System.Drawing.Color.Red);
            }
        }

        private void ShutDown_Click(object sender, EventArgs e)
        {

            OutputBox.Text = "";

            string remotePCName = PCName.Text; // Replace with the actual remote PC's name

            ConnectionOptions options = new ConnectionOptions();
            options.Username = remoteUsername;
            options.Password = remotePassword;

            ManagementScope scope = new ManagementScope($"\\\\{remotePCName}\\root\\cimv2", options);

            try
            {
                scope.Connect();

                ObjectQuery query = new ObjectQuery("SELECT * FROM Win32_OperatingSystem");
                ManagementObjectSearcher searcher = new ManagementObjectSearcher(scope, query);

                ManagementObjectCollection queryCollection = searcher.Get();
                foreach (ManagementObject m in queryCollection)
                {
                    string operatingSystem = m["Caption"].ToString().ToLower();
                    if (operatingSystem.Contains("windows"))
                    {
                        // Initiate the remote shutdown
                        ManagementBaseObject outParams = (ManagementBaseObject)m.InvokeMethod("Win32Shutdown", new object[] { 12, 0 }); // 12 = Shut down the system
                    }
                }
            }
            catch (Exception ex)
            {
                AddColoredText($" - PC {remotePCName} Has Shut Down", System.Drawing.Color.Green);
            }
        }

        private void LogOff_Click(object sender, EventArgs e)
        {
            OutputBox.Text = "";

            string remotePCName = PCName.Text; // Replace with the actual remote PC's name

            ConnectionOptions options = new ConnectionOptions();
            options.Username = remoteUsername;
            options.Password = remotePassword;

            ManagementScope scope = new ManagementScope($"\\\\{remotePCName}\\root\\cimv2", options);

            try
            {
                scope.Connect();

                ObjectQuery query = new ObjectQuery("SELECT * FROM Win32_OperatingSystem");
                ManagementObjectSearcher searcher = new ManagementObjectSearcher(scope, query);

                ManagementObjectCollection queryCollection = searcher.Get();
                foreach (ManagementObject m in queryCollection)
                {
                    string operatingSystem = m["Caption"].ToString().ToLower();
                    if (operatingSystem.Contains("windows"))
                    {
                        // Initiate the remote logoff
                        ManagementBaseObject outParams = (ManagementBaseObject)m.InvokeMethod("Win32Shutdown", new object[] { 4, 0 }); // 4 = Log off the user
                    }
                }
            }
            catch (Exception ex)
            {
                AddColoredText($" - {remotePCName} Has Been Logged Off", System.Drawing.Color.Green);
            }
        }

        private void Restart_Click(object sender, EventArgs e)
        {
            OutputBox.Text = "";

            string remotePCName = PCName.Text; // Replace with the actual remote PC's name

            ConnectionOptions options = new ConnectionOptions();
            options.Username = remoteUsername;
            options.Password = remotePassword;

            ManagementScope scope = new ManagementScope($"\\\\{remotePCName}\\root\\cimv2", options);

            try
            {
                scope.Connect();

                ObjectQuery query = new ObjectQuery("SELECT * FROM Win32_OperatingSystem");
                ManagementObjectSearcher searcher = new ManagementObjectSearcher(scope, query);

                ManagementObjectCollection queryCollection = searcher.Get();
                foreach (ManagementObject m in queryCollection)
                {
                    string operatingSystem = m["Caption"].ToString().ToLower();
                    if (operatingSystem.Contains("windows"))
                    {
                        // Initiate the remote restart
                        ManagementBaseObject outParams = (ManagementBaseObject)m.InvokeMethod("Win32Shutdown", new object[] { 6, 0 }); // 6 = Restart the system
                    }
                }
            }
            catch (Exception ex)
            {
                AddColoredText($" - {remotePCName} Has Been Restarted", System.Drawing.Color.Green);
            }
        }

        private void PCName_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                // Call the function to retrieve remote PC information here                
                RetrieveRemotePCInformation();
                ActiveControl = null;
            }
        }

        private void AddColoredText(string text, System.Drawing.Color color)
        {
            OutputBox.SelectionColor = color;
            OutputBox.AppendText(text);
            OutputBox.SelectionColor = OutputBox.ForeColor; // Reset to the default color
        }

        private void RetrieveRemotePCInformation()
        {
            string remotePCName = PCName.Text; // Replace with the actual remote PC's name
            string UpdatedloggedinUser = User.Text;

            OutputBox.Text = "";

            ConnectionOptions options = new ConnectionOptions();
            options.Username = remoteUsername;
            options.Password = remotePassword;

            ManagementScope scope = new ManagementScope($"\\\\{remotePCName}\\root\\cimv2", options);

            try
            {
                scope.Connect();

                ObjectQuery osQuery = new ObjectQuery("SELECT * FROM Win32_OperatingSystem");
                ObjectQuery memoryQuery = new ObjectQuery("SELECT * FROM Win32_ComputerSystem");
                ObjectQuery cpuQuery = new ObjectQuery("SELECT * FROM Win32_Processor");
                ObjectQuery diskQuery = new ObjectQuery("SELECT * FROM Win32_LogicalDisk WHERE DriveType=3");
                ObjectQuery userQuery = new ObjectQuery("SELECT * FROM Win32_ComputerSystem");

                ManagementObjectSearcher osSearcher = new ManagementObjectSearcher(scope, osQuery);
                ManagementObjectSearcher memorySearcher = new ManagementObjectSearcher(scope, memoryQuery);
                ManagementObjectSearcher cpuSearcher = new ManagementObjectSearcher(scope, cpuQuery);
                ManagementObjectSearcher diskSearcher = new ManagementObjectSearcher(scope, diskQuery);
                ManagementObjectSearcher userSearcher = new ManagementObjectSearcher(scope, userQuery);


                ManagementObjectCollection osCollection = osSearcher.Get();
                ManagementObjectCollection memoryCollection = memorySearcher.Get();
                ManagementObjectCollection cpuCollection = cpuSearcher.Get();
                ManagementObjectCollection diskCollection = diskSearcher.Get();
                ManagementObjectCollection userCollection = userSearcher.Get();

                int textBoxIndex = 0;

                foreach (ManagementObject disk in diskCollection)
                {
                    if (textBoxIndex > 1)
                        break;

                    string driveLetter = disk["DeviceID"].ToString();
                    string volumeName = disk["VolumeName"].ToString();
                    double totalSizeBytes = Convert.ToDouble(disk["Size"]);
                    double freeSpaceBytes = Convert.ToDouble(disk["FreeSpace"]);

                    double totalSizeGB = totalSizeBytes / (1024 * 1024 * 1024);
                    double freeSpaceGB = freeSpaceBytes / (1024 * 1024 * 1024);

                    // Assuming you have text boxes named Drive0 and Drive1
                    System.Windows.Forms.TextBox driveTextBox = Controls.Find($"Drive{textBoxIndex}", true).FirstOrDefault() as System.Windows.Forms.TextBox;

                    if (driveTextBox != null)
                    {
                        driveTextBox.Text = $"{driveLetter} {freeSpaceGB:F2}GB / {totalSizeGB:F2}GB";
                    }

                    textBoxIndex++; // Increment the index for the next set of text boxes
                }

                if (textBoxIndex == 1)
                {
                    // Assuming you have a text box named Drive1
                    System.Windows.Forms.TextBox drive1TextBox = Controls.Find("Drive1", true).FirstOrDefault() as System.Windows.Forms.TextBox;
                    if (drive1TextBox != null)
                    {
                        drive1TextBox.Text = "-";
                    }
                }

                foreach (ManagementObject os in osCollection)
                {
                    string OSname = os["Caption"].ToString();
                    string OSVersion = os["Version"].ToString();

                    if (OSname.Contains("Microsoft"))
                    {
                        OSname = OSname.Replace("Microsoft ", ""); // Remove the specified part
                    }
                    OS.Text = $"{OSname}";
                    // Build.Text = $"{OSVersion}";

                    string codename = GetWindowsCodename(OSVersion);
                    BuildNumber.Text = $"{codename}";
                }

                foreach (ManagementObject memory in memoryCollection)
                {
                    double totalMemoryBytes = Convert.ToDouble(memory["TotalPhysicalMemory"]);
                    double totalMemoryGB = totalMemoryBytes / (1024 * 1024 * 1024);
                    RAM.Text = $"{totalMemoryGB:F2} GB";
                }

                foreach (ManagementObject cpu in cpuCollection)
                {
                    string cpuName = cpu["Name"].ToString();

                    if (cpuName.Contains("Intel(R) Core(TM) "))
                    {
                        cpuName = cpuName.Replace("Intel(R) Core(TM) ", "");
                    }

                    if (cpuName.Contains("CPU @ 3.00GHz"))
                    {
                        cpuName = cpuName.Replace(" CPU @ 3.00GHz", "");
                    }

                    if (cpuName.Contains("CPU @ 3.30GHz"))
                    {
                        cpuName = cpuName.Replace(" CPU @ 3.30GHz", "");
                    }

                    if (cpuName.Contains(" CPU @ 3.20GHz"))
                    {
                        cpuName = cpuName.Replace(" CPU @ 3.20GHz", "");
                    }

                    if (cpuName.Contains(" CPU @ 3.40GHz"))
                    {
                        cpuName = cpuName.Replace(" CPU @ 3.40GHz", "");
                    }

                    if (cpuName.Contains(" v2 @ 3.50GHz"))
                    {
                        cpuName = cpuName.Replace(" v2 @ 3.50GHz", "");
                    }

                    if (cpuName.Contains(" CPU @ 3.20GHz"))
                    {
                        cpuName = cpuName.Replace(" CPU @ 3.20GHz", "");
                    }

                    if (cpuName.Contains("12th Gen "))
                    {
                        cpuName = cpuName.Replace("12th Gen ", "");
                    }

                    if (cpuName.Contains("0 @ 2.60GHz"))
                    {
                        cpuName = cpuName.Replace("0 @ 2.60GHz", "");
                    }

                    if (cpuName.Contains("Intel(R) Xeon(R) CPU"))
                    {
                        cpuName = cpuName.Replace("Intel(R) Xeon(R) CPU", "Xeon");
                    }

                    if (cpuName.Contains("v3 @ 3.40GHz"))
                    {
                        cpuName = cpuName.Replace("v3 @ 3.40GHz", "");
                    }

                    if (cpuName.Contains(" CPU @ 3.70GHz"))
                    {
                        cpuName = cpuName.Replace(" CPU @ 3.70GHz", "");
                    }

                    if (cpuName.Contains(" CPU @ 2.90GHz"))
                    {
                        cpuName = cpuName.Replace(" CPU @ 2.90GHz", "");
                    }

                    CPU.Text = $"{cpuName}";
                }

                Dictionary<string, string> displayNameMapping = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

                string csvUrl = "https://raw.githubusercontent.com/Banana-Goats/Ableworld/main/Users.csv"; // Provide the correct URL to your CSV file
                using (WebClient client = new WebClient())
                {
                    string csvContent = client.DownloadString(csvUrl);
                    string[] lines = csvContent.Split('\n');

                    foreach (string line in lines.Skip(1)) // Skip the header line
                    {
                        string[] values = line.Split(',');
                        if (values.Length == 2)
                        {
                            string username = values[0].Trim(); // Trim whitespace
                            string displayName = values[1].Trim(); // Trim whitespace
                            displayNameMapping[username] = displayName;
                        }
                    }
                }

                foreach (ManagementObject user in userCollection)
                {
                    try
                    {
                        if (user != null)
                        {
                            string loggedinUser = user["UserName"]?.ToString(); // Use the null-conditional operator (?)

                            // Check if loggedinUser is not null before further processing
                            if (loggedinUser != null)
                            {
                                // Remove domain prefix
                                int backslashIndex = loggedinUser.IndexOf('\\');
                                if (backslashIndex != -1)
                                {
                                    loggedinUser = loggedinUser.Substring(backslashIndex + 1);
                                }

                                // Check if the user is in the dictionary
                                if (displayNameMapping.ContainsKey(loggedinUser))
                                {
                                    loggedinUser = displayNameMapping[loggedinUser];
                                }
                                else
                                {
                                    // Handle the case when the user is not found in the dictionary
                                    // You can display a default name or take other appropriate action.
                                    loggedinUser = "Unknown User"; // Replace with your preferred default value.
                                }

                                // Assuming you have a TextBox named LoggedInUserTextBox
                                if (User != null) // Check if User TextBox is not null
                                {
                                    User.Text = loggedinUser;
                                    AddColoredText($" - Data For ( {remotePCName} - {loggedinUser} ) Has Been Loaded", System.Drawing.Color.Green);
                                }
                                else
                                {
                                    // Handle the case where the User TextBox is null
                                    // You can display an error message or take other appropriate action.
                                    AddColoredText($" - User TextBox is null - How Did You get That To Happern ?", System.Drawing.Color.Red);
                                }
                            }
                            else
                            {
                                // Handle the case where "UserName" is null or not available
                                // You can display an error message or take other appropriate action.
                                AddColoredText($" - {remotePCName} Isnt Logged In", System.Drawing.Color.Red);
                                User.Text = "Not Logged In";
                            }
                        }
                        else
                        {
                            // Handle the case where the "user" object itself is null
                            // You can display an error message or take other appropriate action.
                            AddColoredText($" - Error: 'user' object is null.", System.Drawing.Color.Red);
                        }
                    }
                    catch (Exception ex)
                    {
                        // Handle the exception (e.g., log it, display an error message, etc.)
                        // You can customize this part to suit your error handling needs.
                        AddColoredText($"Error: {ex.Message}", System.Drawing.Color.Red);
                    }
                }

                Dictionary<string, string> usernameToDepartmentMapping = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

                string csvUrl1 = "https://raw.githubusercontent.com/Banana-Goats/Ableworld/main/Department.csv"; // Provide the correct URL to your new CSV file
                using (WebClient client = new WebClient())
                {
                    string csvContent = client.DownloadString(csvUrl1);
                    string[] lines = csvContent.Split('\n');

                    foreach (string line in lines.Skip(1)) // Skip the header line
                    {
                        string[] values = line.Split(',');
                        if (values.Length == 2)
                        {
                            string username = values[0].Trim(); // Trim whitespace
                            string department = values[1].Trim(); // Trim whitespace
                            usernameToDepartmentMapping[username] = department;
                        }
                    }
                }

                // Assuming you have a TextBox named User and another TextBox named DepartmentTextBox
                string usernameToLookup = User.Text.Trim(); // Get the text from the User TextBox and trim whitespace

                // Check if the username is in the dictionary
                if (usernameToDepartmentMapping.ContainsKey(usernameToLookup))
                {
                    // Get the department from the dictionary
                    string department = usernameToDepartmentMapping[usernameToLookup];

                    // Populate the DepartmentTextBox with the department
                    Department.Text = department;
                }
                else
                {
                    // If the username is not found in the dictionary, display an error message or handle it as needed
                    Department.Text = "No Department"; // You can customize this message
                }

            }
            catch (Exception ex)
            {
                OS.Text = string.Empty;
                RAM.Text = string.Empty;
                CPU.Text = string.Empty;
                BuildNumber.Text = string.Empty;
                Drive0.Text = string.Empty;
                Drive1.Text = string.Empty;
                User.Text = string.Empty;
                Department.Text = string.Empty;
                AddColoredText($" - {remotePCName} Appears To Be Offline", System.Drawing.Color.Red);
            }
        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            ShutDown.Enabled = checkBox1.Checked;
            LogOff.Enabled = checkBox1.Checked;
            Restart.Enabled = checkBox1.Checked;
        }

        private void SystemUsage()
        {
            string remoteComputerName = PCName.Text.Trim();
            if (string.IsNullOrEmpty(remoteComputerName))
            {
                // Handle the case where PCName.Text is empty or only contains spaces
                return;
            }

            ConnectionOptions options = new ConnectionOptions
            {
                Username = remoteUsername,
                Password = remotePassword,
                Impersonation = ImpersonationLevel.Impersonate,
                Authentication = AuthenticationLevel.PacketPrivacy
            };

            ManagementScope scope = new ManagementScope($@"\\{remoteComputerName}\root\cimv2", options);

            // Query to get CPU usage
            ObjectQuery cpuQuery = new ObjectQuery("SELECT * FROM Win32_PerfFormattedData_PerfOS_Processor WHERE Name='_Total'");
            ManagementObjectSearcher cpuSearcher = new ManagementObjectSearcher(scope, cpuQuery);

            try
            {
                foreach (ManagementObject queryObj in cpuSearcher.Get())
                {
                    double cpuUsage = Convert.ToDouble(queryObj["PercentProcessorTime"]);
                    UpdateProgressBar(CPUBar, cpuUsage, CPULabel);
                    refreshTimer.Start();
                }
            }
            catch (UnauthorizedAccessException ex)
            {
                // Handle unauthorized access exception
                Invoke(new Action(() => OutputBox.Text = string.Empty));
                Invoke(new Action(() => AddColoredText(" - Account Isnt A Domain Admin, Please Use Domain Admin Details", System.Drawing.Color.Red)));
                refreshTimer.Stop();

                return; // or handle it in another way
            }
            catch (ArgumentException argEx)
            {
                // Authentication details are blank
                Invoke(new Action(() => OutputBox.Text = string.Empty));
                Invoke(new Action(() => AddColoredText(" - Empty Login Fields, Please Update Them", System.Drawing.Color.Red)));
                refreshTimer.Stop();

                return; // or handle it in another way
            }

            // Query to get RAM usage
            ObjectQuery ramQuery = new ObjectQuery("SELECT * FROM Win32_OperatingSystem");
            ManagementObjectSearcher ramSearcher = new ManagementObjectSearcher(scope, ramQuery);

            foreach (ManagementObject queryObj in ramSearcher.Get())
            {
                ulong totalMemory = Convert.ToUInt64(queryObj["TotalVisibleMemorySize"]);
                ulong freeMemory = Convert.ToUInt64(queryObj["FreePhysicalMemory"]);

                double ramUsage = ((double)(totalMemory - freeMemory) / totalMemory) * 100.0;
                UpdateProgressBar(RAMBar, ramUsage, RAMLabel);
            }

            ObjectQuery networkQuery = new ObjectQuery("SELECT * FROM Win32_PerfFormattedData_Tcpip_NetworkInterface");
            ManagementObjectSearcher networkSearcher = new ManagementObjectSearcher(scope, networkQuery);

            foreach (ManagementObject queryObj in networkSearcher.Get())
            {
                string interfaceName = queryObj["Name"].ToString();
                ulong bytesReceived = Convert.ToUInt64(queryObj["BytesReceivedPerSec"]);
                ulong bytesSent = Convert.ToUInt64(queryObj["BytesSentPerSec"]);

                // Calculate bandwidth usage in megabits per second (Mbps) with 2 decimal places
                ulong totalBytesPerSec = bytesReceived + bytesSent;
                double bandwidthUsageMbps = (totalBytesPerSec * 8) / 1_000_000.0; // Convert to Mbps

                // Display the formatted bandwidth usage for the first interface
                UpdateBandwidthUsage(interfaceName, bandwidthUsageMbps);

                // Exit the loop after processing the first network adapter
                break;
            }
        }

        private void UpdateBandwidthUsage(string interfaceName, double usageMbps)
        {
            if (interfaceTextBox.InvokeRequired)
            {
                interfaceTextBox.Invoke(new Action(() => interfaceTextBox.Text = interfaceName));
            }
            else
            {
                interfaceTextBox.Text = interfaceName;
            }

            if (speedTextBox.InvokeRequired)
            {
                speedTextBox.Invoke(new Action(() => speedTextBox.Text = $"{usageMbps:0.00} Mbps"));
            }
            else
            {
                speedTextBox.Text = $"{usageMbps:0.00} Mbps";
            }
        }


        private void UpdateProgressBar(System.Windows.Forms.ProgressBar progressBar, double usage, Label label)
        {
            if (progressBar.InvokeRequired)
            {
                progressBar.Invoke(new Action(() => UpdateProgressBar(progressBar, usage, label)));
            }
            else
            {
                progressBar.Value = Math.Min((int)usage, progressBar.Maximum);
                label.Text = $"{usage:0.00}%"; // Update the label with the percentage value
            }
        }
        private void TabControl1_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (tabControl1.SelectedTab == Information)
            {
                refreshTimer.Start();
            }
            else
            {
                refreshTimer.Stop();
            }
        }

        private void Export_Click(object sender, EventArgs e)
        {
            // Define the CSV file path
            string csvFilePath = "\\\\able-fs03\\IT Software\\Scripts\\Checklist\\PC's.csv";

            // Read existing CSV data
            List<string[]> csvData = new List<string[]>();

            if (File.Exists(csvFilePath))
            {
                using (StreamReader reader = new StreamReader(csvFilePath))
                {
                    string line;
                    while ((line = reader.ReadLine()) != null)
                    {
                        string[] values = line.Split(',');
                        csvData.Add(values);
                    }
                }
            }

            // Get the PCName directly from your source (e.g., a variable or input)
            string pcName = PCName.Text; // Replace with your source of PCName

            // Find the record with the specific PCName
            bool found = false;
            for (int i = 0; i < csvData.Count; i++)
            {
                if (csvData[i][0] == pcName)
                {
                    // Update the data
                    csvData[i][1] = Department.Text;
                    csvData[i][2] = User.Text;
                    csvData[i][3] = OS.Text;
                    csvData[i][4] = BuildNumber.Text;
                    csvData[i][5] = RAM.Text;
                    csvData[i][6] = CPU.Text;
                    csvData[i][7] = Drive0.Text;
                    csvData[i][8] = Drive1.Text;
                    found = true;
                    break;
                }
            }

            if (!found)
            {
                // If the PCName doesn't exist, create a new record
                string[] newRecord = new string[]
                {
                    pcName,
                    Department.Text,
                    User.Text,
                    OS.Text,
                    BuildNumber.Text,
                    RAM.Text,
                    CPU.Text,
                    Drive0.Text,
                    Drive1.Text
                };
                csvData.Add(newRecord);
            }

            // Write the updated data (including new records) back to the CSV file
            using (StreamWriter writer = new StreamWriter(csvFilePath))
            {
                foreach (var values in csvData)
                {
                    writer.WriteLine(string.Join(",", values));
                }
            }
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            tabControl1.SelectedTab = About;
        }
    }
}







