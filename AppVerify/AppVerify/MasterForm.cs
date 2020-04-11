using System;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Management;
using System.Diagnostics;
using System.IO;
using Microsoft.Win32;
using System.Text.RegularExpressions;
using System.Net;
using System.Reflection;

namespace AppVerify
{

    public partial class MasterForm : Form
    {
        private String BiosData;
        private String MAKData;

        public MasterForm()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void ExecuteCommand(String command)
        {
            ProcessStartInfo ProcessInfo;
            ProcessInfo = new ProcessStartInfo("cmd.exe", "/k" + command);
            ProcessInfo.Verb = "runas"; // Run CMD as admin
            ProcessInfo.CreateNoWindow = true;
            ProcessInfo.UseShellExecute = false;
            ProcessInfo.RedirectStandardOutput = true;

            Process.Start(ProcessInfo);
        }

        private void ExecuteOUTPUTCommand(String command)
        {
            ProcessStartInfo ProcessInfo;
            ProcessInfo = new ProcessStartInfo("cmd.exe", "/k" + command);
            ProcessInfo.Verb = "runas"; // Run CMD as admin
            ProcessInfo.CreateNoWindow = true;
            ProcessInfo.UseShellExecute = false;
            ProcessInfo.RedirectStandardOutput = true;

            Process.Start(ProcessInfo);
            System.Threading.Thread.Sleep(50000);

        }

        /**private string DOELogo()
        {
            string value = " ";
            ProcessStartInfo ProcessInfo;
            ProcessInfo = new ProcessStartInfo("cmd.exe", "/k" + "control /name microsoft.system");
            ProcessInfo.CreateNoWindow = true;
            ProcessInfo.UseShellExecute = false;
            ProcessInfo.RedirectStandardOutput = false;

            Process.Start(ProcessInfo);

            System.Threading.Thread.Sleep(3000);
            DialogResult dialogResult = MessageBox.Show(new Form { TopMost = true }, "Do you see the DOE Logo?", "DOE Logo Verification", MessageBoxButtons.YesNo);
            if (dialogResult == DialogResult.Yes)
            {
                value = "YES";
            }
            else if (dialogResult == DialogResult.No)
            {
                value = "NO";
            }

            return value;
        }**/

        private void ExecuteProgramsCommand(string command)
        {
            ProcessStartInfo StartInfo;
            StartInfo = new ProcessStartInfo("cmd.exe", "/k" + command);
            StartInfo.Verb = "runas"; // Run CMD as admin
            StartInfo.CreateNoWindow = true;
            StartInfo.UseShellExecute = false;
            StartInfo.RedirectStandardOutput = true;

            Process.Start(StartInfo);
        } // For Testing of Installed Programs

        private String getVersionAutomatically(string version)
        {
            string CommandInstalledApps = @"wmic product get name, version >C:\allprograms.txt";
            ExecuteProgramsCommand(CommandInstalledApps);

            System.Threading.Thread.Sleep(40000);

            using (StreamReader reader = File.OpenText(@"c:\allprograms.txt"))
            {
                string currentLine;
                while ((currentLine = reader.ReadLine()) != null)
                {

                    if (currentLine.Contains("Silverlight"))
                    {
                        version = currentLine;

                        string pattern = "\\s+";
                        string replacement = "";
                        Regex rgx = new Regex(pattern);
                        version = rgx.Replace(version, replacement);

                        pattern = "[A-Za-z]";
                        replacement = "";
                        rgx = new Regex(pattern);
                        version = rgx.Replace(version, replacement);
                    }

                }

                reader.Close();
                File.Delete(@"c:\allprograms.txt");
            }

            return version;
        }

        public void KillCommandProcesses()
        {
            System.Diagnostics.Process[] AppProcesses = System.Diagnostics.Process.GetProcessesByName("APPVERIFY");

            foreach (System.Diagnostics.Process p in AppProcesses)
            {
                if (!string.IsNullOrEmpty(p.ProcessName) && p.ProcessName != "AppVerify")
                {
                    try
                    {
                        p.Kill();
                    }
                    catch
                    {

                    }
                }
            }
        }

        private void RunChecklist_Click(object sender, EventArgs e)
        {
            System.Diagnostics.Process[] ExcelProcesses = System.Diagnostics.Process.GetProcessesByName("EXCEL");

            foreach (System.Diagnostics.Process p in ExcelProcesses)
            {
                if (!string.IsNullOrEmpty(p.ProcessName))
                {
                    try
                    {
                        p.Kill();
                    }
                    catch
                    {

                    }
                }
            } // Checks if AppVerify.csv is already running on the computer and kills the process

            System.Threading.Thread.Sleep(2000);

            if (File.Exists(@"C:\AppVerify.csv"))
            {
                File.Delete(@"C:\AppVerify.csv");
            } // Deletes the AppVerify.csv file if already on computer

            MessageBox.Show("AppVerify will start running after this messagebox is closed. Please do not close the application or open any other programs. Prompts will appear soon to provide further instructions. If applications displays 'Not Responding' please allow it to complete it's processing. ", "AppVerify Alert");

            string Image_Information = "Could not be found, File may not exist";

            if (File.Exists(@"C:\DOE INSTALLS\Image_Information.htm"))
            {
                try
                {
                    Process.Start(@"C:\DOE INSTALLS\Image_Information.htm");
                    Process.Start(@"C:\Windows\System32");
                }
                catch
                {

                }
                DialogResult ImageResult = MessageBox.Show(new Form { TopMost = true }, "Is Image Version file in System32? Please read the Image_Information file and find the version # under Image Version. Verify that this value is the same as the one in the System32 folder by matching the values. Use the search function in File explorer and type: " + "vL",
                    "Image_Information Verification", MessageBoxButtons.YesNoCancel);
                if (ImageResult == DialogResult.Yes)
                {
                    Image_Information = "YES";
                }
                else if (ImageResult == DialogResult.No)
                {
                    Image_Information = "NO";
                }
                else if (ImageResult == DialogResult.Cancel)
                {
                    Image_Information = "Not Applicable";
                }

            } // Opens Image Information for School Laptops

            else if (File.Exists(@"C:\Drivers\InstalledDrivers.htm"))
            {
                try
                {
                    Process.Start(@"C:\Drivers\InstalledDrivers.htm");
                    Process.Start(@"C:\Windows\System32");
                }
                catch
                {

                }
                DialogResult ImageResult = MessageBox.Show(new Form { TopMost = true }, "Is Image Version file in System32? Please read the Image_Information file and find the version # under Image Version. Verify that this value is the same as the one in the System32 folder by matching the values. Use the search function in File explorer and type: " + "vL",
                    "Image_Information Verification", MessageBoxButtons.YesNoCancel);
                if (ImageResult == DialogResult.Yes)
                {
                    Image_Information = "YES";
                }
                else if (ImageResult == DialogResult.No)
                {
                    Image_Information = "NO";
                }
                else if (ImageResult == DialogResult.Cancel)
                {
                    Image_Information = "Not Applicable";
                }

            } // Opens Image Information File for Central Computers

            System.Threading.Thread.Sleep(2000);

            using (ManagementObjectSearcher mos = new ManagementObjectSearcher("SELECT * FROM Win32_BIOS"))

            using (ManagementObjectCollection mocDriver = mos.Get())

            {

                StringBuilder sb = new StringBuilder();
                foreach (ManagementObject mo in mocDriver)

                {

                    string[] BIOSVersions = (string[])mo["BIOSVersion"];

                    sb.Append("BIOS Manufacturer: " + mo["Manufacturer"].ToString() + ", ");
                    sb.Append("SMBIOSBIOS Version: " + mo["SMBIOSBIOSVersion"].ToString() + ", ");

                    string temp = mo["ReleaseDate"].ToString();
                    temp = temp.Substring(0, 8); // Parses the DATE information

                    DateTime Biosdate = DateTime.ParseExact(temp, "yyyyMMdd", null);
                    temp = Biosdate.ToString().Substring(0, 9);
                    sb.Append("BIOS Release Date: " + temp);
                    BiosData = sb.ToString(); 
                }

            } // Gets Bios Information

            ExecuteCommand(@"cscript C:\Windows\System32\slmgr.vbs /dli >C:\AppVerifyMAK.txt");

            System.Threading.Thread.Sleep(8000); // application stops for 8 seconds

            string pattern = "Partial";
            string pattern2 = "License";
            string pattern3 = "Description";
            StringBuilder MAK = new StringBuilder();

            using (StreamReader reader = File.OpenText(@"C:\AppVerifyMAK.txt"))
            {
                string currentLine;
                while ((currentLine = reader.ReadLine()) != null)
                {
                    if (currentLine.Contains(pattern))
                    {
                        MAK.Append(currentLine + ", ");
                    }

                    if (currentLine.Contains(pattern2))
                    {
                        MAK.Append(currentLine);
                    }

                    if (currentLine.Contains(pattern3))
                    {
                        MAK.Append(currentLine + ",");
                    }

                }

                reader.Close();
                File.Delete(@"C:\AppVerifyMAK.txt");

            } // Mak Information, Windows License Information

            MAKData = MAK.ToString();

            ExecuteOUTPUTCommand(@"cscript ""C:\Program Files (x86)\Microsoft Office\Office16\OSPP.VBS"" /dstatus >C:\AppVerifyOFFICE.txt");
            // Office 2016 Information
            System.Threading.Thread.Sleep(10000); // application stops for 10 seconds

            string HostName = Dns.GetHostName();
            string IP = Dns.GetHostByName(HostName).AddressList[0].ToString();
            //IP Information
            
            if (IP.Equals("127.0.0.1") || IP.Equals("0.0.0.0"))
            {
                IP = "IP_Address: " + "No IP_Address could be found. Please check connection";
            }

            pattern = "LICENSE NAME";
            pattern2 = "LICENSE STATUS";
            pattern3 = "Last 5";
            StringBuilder OFFICE = new StringBuilder();

            using (StreamReader reader = File.OpenText(@"C:\AppVerifyOFFICE.txt"))
            {
                string currentLine;
                while ((currentLine = reader.ReadLine()) != null)
                {
                    if (currentLine.Contains(pattern))
                    {
                        OFFICE.Append(currentLine + ", ");
                    }

                    if (currentLine.Contains(pattern2))
                    {
                        OFFICE.Append(currentLine + ", ");
                    }

                    if (currentLine.Contains(pattern3))
                    {
                        OFFICE.Append(currentLine + ", ");
                    }
                }

                reader.Close();
                File.Delete(@"C:\AppVerifyOFFICE.txt");
            } // Office 2016 Information

            string OfficeData = OFFICE.ToString();

            string DOEinstalls = "";
            string DOEinstallsVersion = "";

            try
            {
                if (Directory.Exists(@"C:\DOE INSTALLS")) {
                    DOEinstalls = "YES";

                }
                else
                {
                    DOEinstalls = "NO";
                }

            }
            catch
            {

            }

            //string value = DOELogo();

            if (Directory.Exists(@"C:\DOE INSTALLS\Winstaller-Images"))
            {
                //MessageBox.Show(Directory.GetCurrentDirectory() + @"\CertUtil.exe");
                if (File.Exists(@"C:\DOE INSTALLS\Winstaller-Images\Ver 1.3"))
                {
                    DOEinstallsVersion = "YES";
                }

                else
                {
                    DOEinstallsVersion = "NO";
                }
            }

            else
            {
                DOEinstallsVersion = "DOE INSTALLS FOLDER was not found, VER 1.3 cannot be confirmed!";
            }


            //Power VERIFICATION
            Boolean SchoolLaptop = false;
            Boolean CentralDesktop = false;
            Boolean SchoolDesktop = false;


            DialogResult PowerDialog1 = MessageBox.Show(new Form { TopMost = true }, "Is this a School Laptop or Tablet?", "Power Verification", MessageBoxButtons.YesNo);
            if (PowerDialog1 == DialogResult.Yes)
            {
                SchoolLaptop = true;
            }
            if (PowerDialog1 == DialogResult.No)
            {
                DialogResult PowerDialog2 = MessageBox.Show(new Form { TopMost = true }, "Is this a School Desktop?", "Power Verification", MessageBoxButtons.YesNo);
                if (PowerDialog2 == DialogResult.Yes)
                {
                    SchoolDesktop = true;
                }
                if (PowerDialog2 == DialogResult.No)
                {
                    DialogResult PowerDialog3 = MessageBox.Show(new Form { TopMost = true }, "Is this a Central Desktop?", "Power Verification", MessageBoxButtons.YesNo);
                    if (PowerDialog3 == DialogResult.Yes)
                    {
                        CentralDesktop = true;
                    }
                    if (PowerDialog3 == DialogResult.No)
                    {
                        CentralDesktop = false;
                    }
                }
            }

            System.Threading.Thread.Sleep(5000);

            //AC IS PLUGGED IN
            //DC IS RUNNING ON BATTERIES

            string HardDisk = "";
            string SleepSettings = "";
            string HibernateSettings = "";
            string LidClose = "";
            string PowerButton = "";
            string Display = "";
            int tempPower = 0;

            if (CentralDesktop)
            {
                ExecuteCommand(@"powercfg -query 0ec330c5-0896-483c-a480-7ac80476a69d 0012ee47-9041-4b5d-9b77-535fba8b1442 6738e2c4-e8a5-4a42-b16a-e040e769756e > c:\powerplanHardDisk.txt");
                System.Threading.Thread.Sleep(2000);

                string HardAC = "";
                string HardDC = "";

                using (StreamReader reader = File.OpenText(@"c:\powerplanHardDisk.txt"))
                {
                    string currentLine;
                    while ((currentLine = reader.ReadLine()) != null)
                    {
                        if (currentLine.Contains("AC"))
                        {
                            HardAC = currentLine;
                            HardAC = HardAC.Remove(0, 38);
                            tempPower = Int32.Parse(HardAC, System.Globalization.NumberStyles.HexNumber);
                            tempPower = tempPower / 60;
                            HardAC = tempPower.ToString() + " Min.";
                            //MessageBox.Show(SleepAC);

                        }

                        if (currentLine.Contains("DC"))
                        {
                            HardDC = currentLine;
                            HardDC = HardDC.Remove(0, 38); // Parses and removes 0 -> 38 values from the currentLine string
                            tempPower = Int32.Parse(HardDC, System.Globalization.NumberStyles.HexNumber); // Converts from hexadecimal to int
                            tempPower = tempPower / 60; // Converts from seconds to minutes
                            HardDC = tempPower.ToString() + " Min.";
                            //MessageBox.Show(SleepDC);
                        }

                    }

                } // Power Plan Hard Disk Information

                File.Delete(@"c:\powerplanHardDisk.txt");
                HardDisk = HardDC + "," + HardAC;

                ExecuteCommand(@"powercfg -query 0ec330c5-0896-483c-a480-7ac80476a69d 238c9fa8-0aad-41ed-83f4-97be242c8f20 29f6c1db-86da-48c5-9fdb-f2b67b1f44da >c:\powerplanSleep.txt");
                System.Threading.Thread.Sleep(2000);

                string SleepAC = "";
                string SleepDC = "";

                using (StreamReader reader = File.OpenText(@"c:\powerplanSleep.txt"))
                {
                    string currentLine;
                    while ((currentLine = reader.ReadLine()) != null)
                    {
                        if (currentLine.Contains("AC"))
                        {
                            SleepAC = currentLine;
                            SleepAC = SleepAC.Remove(0, 38);
                            tempPower = Int32.Parse(SleepAC, System.Globalization.NumberStyles.HexNumber);
                            tempPower = tempPower / 60;
                            SleepAC = tempPower.ToString() + " Min.";
                            //MessageBox.Show(SleepAC);

                        }

                        if (currentLine.Contains("DC"))
                        {
                            SleepDC = currentLine;
                            SleepDC = SleepDC.Remove(0, 38);
                            tempPower = Int32.Parse(SleepDC, System.Globalization.NumberStyles.HexNumber);
                            tempPower = tempPower / 60;
                            SleepDC = tempPower.ToString() + " Min.";
                            //MessageBox.Show(SleepDC);
                        }

                    }

                }

                File.Delete(@"c:\powerplanSleep.txt");
                SleepSettings = SleepDC + "," + SleepAC;
                
                ExecuteCommand(@"powercfg -query 0ec330c5-0896-483c-a480-7ac80476a69d 238c9fa8-0aad-41ed-83f4-97be242c8f20 9d7815a6-7ee4-497e-8888-515a05f02364 >c:\powerplanHibernate.txt");
                System.Threading.Thread.Sleep(2000);

                string HibernateAC = "";
                string HibernateDC = "";

                using (StreamReader reader = File.OpenText(@"c:\powerplanHibernate.txt"))
                {
                    string currentLine;
                    while ((currentLine = reader.ReadLine()) != null)
                    {
                        if (currentLine.Contains("AC"))
                        {
                            HibernateAC = currentLine;
                            HibernateAC = HibernateAC.Remove(0, 38);
                            tempPower = Int32.Parse(HibernateAC, System.Globalization.NumberStyles.HexNumber);
                            tempPower = tempPower / 60;
                            HibernateAC = tempPower.ToString() + " Min.";
                            //MessageBox.Show(HibernateAC);

                        }

                        if (currentLine.Contains("DC"))
                        {
                            HibernateDC = currentLine;
                            HibernateDC = HibernateDC.Remove(0, 38);
                            tempPower = Int32.Parse(HibernateDC, System.Globalization.NumberStyles.HexNumber);
                            tempPower = tempPower / 60;
                            HibernateDC = tempPower.ToString() + " Min.";
                            //MessageBox.Show(HibernateDC);
                        }
                    }
                }

                File.Delete(@"c:\powerplanHibernate.txt");
                HibernateSettings = HibernateDC + "," + HibernateAC;

                ExecuteCommand(@"powercfg -query 0ec330c5-0896-483c-a480-7ac80476a69d 4f971e89-eebd-4455-a8de-9e59040e7347 5ca83367-6e45-459f-a27b-476b1d01c936 >c:\powerplanLidClose.txt");
                System.Threading.Thread.Sleep(2000);
                string LidAC = "";
                string LidDC = "";

                using (StreamReader reader = File.OpenText(@"c:\powerplanLidClose.txt"))
                {
                    string currentLine;
                    while ((currentLine = reader.ReadLine()) != null)
                    {
                        if (currentLine.Contains("AC"))
                        {
                            LidAC = currentLine;
                            LidAC = LidAC.Remove(0, LidAC.Length - 3);
                            if (LidAC.Equals("000"))
                            {
                                LidAC = "Do nothing";
                            }
                            if (LidAC.Equals("001"))
                            {
                                LidAC = "Sleep";
                            }
                            if (LidAC.Equals("002"))
                            {
                                LidAC = "Hibernate";
                            }
                            if (LidAC.Equals("003"))
                            {
                                LidAC = "Shut down";
                            }
                        }

                        if (currentLine.Contains("DC"))
                        {
                            LidDC = currentLine;
                            LidDC = LidDC.Remove(0, LidDC.Length - 3);
                            if (LidDC.Equals("000"))
                            {
                                LidDC = "Do nothing";
                            }
                            if (LidDC.Equals("001"))
                            {
                                LidDC = "Sleep";
                            }
                            if (LidDC.Equals("002"))
                            {
                                LidDC = "Hibernate";
                            }
                            if (LidDC.Equals("003"))
                            {
                                LidDC = "Shut down";
                            }
                        }
                    }
                }

                File.Delete(@"c:\powerplanLidClose.txt");
                LidClose = LidDC + "," + LidAC;

                ExecuteCommand(@"powercfg -query 0ec330c5-0896-483c-a480-7ac80476a69d 4f971e89-eebd-4455-a8de-9e59040e7347 7648efa3-dd9c-4e3e-b566-50f929386280  >c:\powerplanButtonAction.txt");
                System.Threading.Thread.Sleep(2000);
                string PowerButtonAC = "";
                string PowerButtonDC = "";

                using (StreamReader reader = File.OpenText(@"c:\powerplanButtonAction.txt"))
                {
                    string currentLine;
                    while ((currentLine = reader.ReadLine()) != null)
                    {
                        if (currentLine.Contains("AC"))
                        {
                            PowerButtonAC = currentLine;
                            PowerButtonAC = PowerButtonAC.Remove(0, PowerButtonAC.Length - 3);
                            if (PowerButtonAC.Equals("000"))
                            {
                                PowerButtonAC = "Do nothing";
                            }
                            if (PowerButtonAC.Equals("001"))
                            {
                                PowerButtonAC = "Sleep";
                            }
                            if (PowerButtonAC.Equals("002"))
                            {
                                PowerButtonAC = "Hibernate";
                            }
                            if (PowerButtonAC.Equals("003"))
                            {
                                PowerButtonAC = "Shut down";
                            }
                            if (PowerButtonAC.Equals("004"))
                            {
                                PowerButtonAC = "Turn off display";
                            }

                        }

                        if (currentLine.Contains("DC"))
                        {
                            PowerButtonDC = currentLine;
                            PowerButtonDC = PowerButtonDC.Remove(0, PowerButtonDC.Length - 3);
                            if (PowerButtonDC.Equals("000"))
                            {
                                PowerButtonDC = "Do nothing";
                            }
                            if (PowerButtonDC.Equals("001"))
                            {
                                PowerButtonDC = "Sleep";
                            }
                            if (PowerButtonDC.Equals("002"))
                            {
                                PowerButtonDC = "Hibernate";
                            }
                            if (PowerButtonDC.Equals("003"))
                            {
                                PowerButtonDC = "Shut down";
                            }
                            if (PowerButtonDC.Equals("004"))
                            {
                                PowerButtonDC = "Turn off display";
                            }

                        }
                    }
                }

                File.Delete(@"c:\powerplanButtonAction.txt");
                PowerButton = PowerButtonDC + "," + PowerButtonAC;

                ExecuteCommand(@"powercfg -query 0ec330c5-0896-483c-a480-7ac80476a69d 7516b95f-f776-4464-8c53-06167f40cc99 3c0bc021-c8a8-4e07-a973-6b14cbcb2b7e  >c:\powerplanDisplay.txt");
                System.Threading.Thread.Sleep(2000);

                string DisplayAC = "";
                string DisplayDC = "";

                using (StreamReader reader = File.OpenText(@"c:\powerplanDisplay.txt"))
                {
                    string currentLine;
                    while ((currentLine = reader.ReadLine()) != null)
                    {
                        if (currentLine.Contains("AC"))
                        {
                            DisplayAC = currentLine;
                            DisplayAC = DisplayAC.Remove(0, 38);
                            tempPower = Int32.Parse(DisplayAC, System.Globalization.NumberStyles.HexNumber);
                            tempPower = tempPower / 60;
                            DisplayAC = tempPower.ToString() + " Min.";
                            //MessageBox.Show(HibernateAC);

                        }

                        if (currentLine.Contains("DC"))
                        {
                            DisplayDC = currentLine;
                            DisplayDC = DisplayDC.Remove(0, 38);
                            tempPower = Int32.Parse(DisplayDC, System.Globalization.NumberStyles.HexNumber);
                            tempPower = tempPower / 60;
                            DisplayDC = tempPower.ToString() + " Min.";
                            //MessageBox.Show(HibernateDC);
                        }
                    }
                }

                File.Delete(@"c:\powerplanDisplay.txt");
                Display = DisplayDC + "," + DisplayAC;
            }


            if (SchoolLaptop)
            {
                ExecuteCommand(@"powercfg -query 8b9af922-23ef-4a61-8f14-f974e479ca1a 0012ee47-9041-4b5d-9b77-535fba8b1442 6738e2c4-e8a5-4a42-b16a-e040e769756e  > c:\powerplanHardDisk.txt");
                System.Threading.Thread.Sleep(2000);

                string HardAC = "";
                string HardDC = "";

                using (StreamReader reader = File.OpenText(@"c:\powerplanHardDisk.txt"))
                {
                    string currentLine;
                    while ((currentLine = reader.ReadLine()) != null)
                    {
                        if (currentLine.Contains("AC"))
                        {
                            HardAC = currentLine;
                            HardAC = HardAC.Remove(0, 38);
                            tempPower = Int32.Parse(HardAC, System.Globalization.NumberStyles.HexNumber);
                            tempPower = tempPower / 60;
                            HardAC = tempPower.ToString() + " Min.";
                            //MessageBox.Show(SleepAC);

                        }

                        if (currentLine.Contains("DC"))
                        {
                            HardDC = currentLine;
                            HardDC = HardDC.Remove(0, 38);
                            tempPower = Int32.Parse(HardDC, System.Globalization.NumberStyles.HexNumber);
                            tempPower = tempPower / 60;
                            HardDC = tempPower.ToString() + " Min.";
                            //MessageBox.Show(SleepDC);
                        }

                    }

                }

                File.Delete(@"c:\powerplanHardDisk.txt");
                HardDisk = HardDC + "," + HardAC;

                ExecuteCommand(@"powercfg -query 8b9af922-23ef-4a61-8f14-f974e479ca1a 238c9fa8-0aad-41ed-83f4-97be242c8f20 29f6c1db-86da-48c5-9fdb-f2b67b1f44da  >c:\powerplanSleep.txt");
                System.Threading.Thread.Sleep(2000);

                string SleepAC = "";
                string SleepDC = "";

                using (StreamReader reader = File.OpenText(@"c:\powerplanSleep.txt"))
                {
                    string currentLine;
                    while ((currentLine = reader.ReadLine()) != null)
                    {
                        if (currentLine.Contains("AC"))
                        {
                            SleepAC = currentLine;
                            SleepAC = SleepAC.Remove(0, 38);
                            tempPower = Int32.Parse(SleepAC, System.Globalization.NumberStyles.HexNumber);
                            tempPower = tempPower / 60;
                            SleepAC = tempPower.ToString() + " Min.";
                            //MessageBox.Show(SleepAC);

                        }

                        if (currentLine.Contains("DC"))
                        {
                            SleepDC = currentLine;
                            SleepDC = SleepDC.Remove(0, 38);
                            tempPower = Int32.Parse(SleepDC, System.Globalization.NumberStyles.HexNumber);
                            tempPower = tempPower / 60;
                            SleepDC = tempPower.ToString() + " Min.";
                            //MessageBox.Show(SleepDC);
                        }

                    }

                }

                File.Delete(@"c:\powerplanSleep.txt");
                SleepSettings = SleepDC + "," + SleepAC;

                ExecuteCommand(@"powercfg -query 8b9af922-23ef-4a61-8f14-f974e479ca1a 238c9fa8-0aad-41ed-83f4-97be242c8f20 9d7815a6-7ee4-497e-8888-515a05f02364 >c:\powerplanHibernate.txt");
                System.Threading.Thread.Sleep(2000);

                string HibernateAC = "";
                string HibernateDC = "";

                using (StreamReader reader = File.OpenText(@"c:\powerplanHibernate.txt"))
                {
                    string currentLine;
                    while ((currentLine = reader.ReadLine()) != null)
                    {
                        if (currentLine.Contains("AC"))
                        {
                            HibernateAC = currentLine;
                            HibernateAC = HibernateAC.Remove(0, 38);
                            tempPower = Int32.Parse(HibernateAC, System.Globalization.NumberStyles.HexNumber);
                            tempPower = tempPower / 60;
                            HibernateAC = tempPower.ToString() + " Min.";
                            //MessageBox.Show(HibernateAC);

                        }

                        if (currentLine.Contains("DC"))
                        {
                            HibernateDC = currentLine;
                            HibernateDC = HibernateDC.Remove(0, 38);
                            tempPower = Int32.Parse(HibernateDC, System.Globalization.NumberStyles.HexNumber);
                            tempPower = tempPower / 60;
                            HibernateDC = tempPower.ToString() + " Min.";
                            //MessageBox.Show(HibernateDC);
                        }
                    }
                }

                File.Delete(@"c:\powerplanHibernate.txt");
                HibernateSettings = HibernateDC + "," + HibernateAC;

                ExecuteCommand(@"powercfg -query 8b9af922-23ef-4a61-8f14-f974e479ca1a 4f971e89-eebd-4455-a8de-9e59040e7347 5ca83367-6e45-459f-a27b-476b1d01c936  >c:\powerplanLidClose.txt");
                System.Threading.Thread.Sleep(2000);
                string LidAC = "";
                string LidDC = "";

                using (StreamReader reader = File.OpenText(@"c:\powerplanLidClose.txt"))
                {
                    string currentLine;
                    while ((currentLine = reader.ReadLine()) != null)
                    {
                        if (currentLine.Contains("AC"))
                        {
                            LidAC = currentLine;
                            LidAC = LidAC.Remove(0, LidAC.Length - 3);
                            if (LidAC.Equals("000"))
                            {
                                LidAC = "Do nothing";
                            }
                            if (LidAC.Equals("001"))
                            {
                                LidAC = "Sleep";
                            }
                            if (LidAC.Equals("002"))
                            {
                                LidAC = "Hibernate";
                            }
                            if (LidAC.Equals("003"))
                            {
                                LidAC = "Shut down";
                            }
                        }

                        if (currentLine.Contains("DC"))
                        {
                            LidDC = currentLine;
                            LidDC = LidDC.Remove(0, LidDC.Length - 3);
                            if (LidDC.Equals("000"))
                            {
                                LidDC = "Do nothing";
                            }
                            if (LidDC.Equals("001"))
                            {
                                LidDC = "Sleep";
                            }
                            if (LidDC.Equals("002"))
                            {
                                LidDC = "Hibernate";
                            }
                            if (LidDC.Equals("003"))
                            {
                                LidDC = "Shut down";
                            }
                        }
                    }
                }

                File.Delete(@"c:\powerplanLidClose.txt");
                LidClose = LidDC + "," + LidAC;

                ExecuteCommand(@"powercfg -query 8b9af922-23ef-4a61-8f14-f974e479ca1a 4f971e89-eebd-4455-a8de-9e59040e7347 7648efa3-dd9c-4e3e-b566-50f929386280 >c:\powerplanButtonAction.txt");
                System.Threading.Thread.Sleep(2000);
                string PowerButtonAC = "";
                string PowerButtonDC = "";

                using (StreamReader reader = File.OpenText(@"c:\powerplanButtonAction.txt"))
                {
                    string currentLine;
                    while ((currentLine = reader.ReadLine()) != null)
                    {
                        if (currentLine.Contains("AC"))
                        {
                            PowerButtonAC = currentLine;
                            PowerButtonAC = PowerButtonAC.Remove(0, PowerButtonAC.Length - 3);
                            if (PowerButtonAC.Equals("000"))
                            {
                                PowerButtonAC = "Do nothing";
                            }
                            if (PowerButtonAC.Equals("001"))
                            {
                                PowerButtonAC = "Sleep";
                            }
                            if (PowerButtonAC.Equals("002"))
                            {
                                PowerButtonAC = "Hibernate";
                            }
                            if (PowerButtonAC.Equals("003"))
                            {
                                PowerButtonAC = "Shut down";
                            }
                            if (PowerButtonAC.Equals("004"))
                            {
                                PowerButtonAC = "Turn off display";
                            }

                        }

                        if (currentLine.Contains("DC"))
                        {
                            PowerButtonDC = currentLine;
                            PowerButtonDC = PowerButtonDC.Remove(0, PowerButtonDC.Length - 3);
                            if (PowerButtonDC.Equals("000"))
                            {
                                PowerButtonDC = "Do nothing";
                            }
                            if (PowerButtonDC.Equals("001"))
                            {
                                PowerButtonDC = "Sleep";
                            }
                            if (PowerButtonDC.Equals("002"))
                            {
                                PowerButtonDC = "Hibernate";
                            }
                            if (PowerButtonDC.Equals("003"))
                            {
                                PowerButtonDC = "Shut down";
                            }
                            if (PowerButtonDC.Equals("004"))
                            {
                                PowerButtonDC = "Turn off display";
                            }

                        }
                    }
                }

                File.Delete(@"c:\powerplanButtonAction.txt");
                PowerButton = PowerButtonDC + "," + PowerButtonAC;

                ExecuteCommand(@"powercfg -query 8b9af922-23ef-4a61-8f14-f974e479ca1a 7516b95f-f776-4464-8c53-06167f40cc99 3c0bc021-c8a8-4e07-a973-6b14cbcb2b7e   >c:\powerplanDisplay.txt");
                System.Threading.Thread.Sleep(2000);

                string DisplayAC = "";
                string DisplayDC = "";

                using (StreamReader reader = File.OpenText(@"c:\powerplanDisplay.txt"))
                {
                    string currentLine;
                    while ((currentLine = reader.ReadLine()) != null)
                    {
                        if (currentLine.Contains("AC"))
                        {
                            DisplayAC = currentLine;
                            DisplayAC = DisplayAC.Remove(0, 38);
                            tempPower = Int32.Parse(DisplayAC, System.Globalization.NumberStyles.HexNumber);
                            tempPower = tempPower / 60;
                            DisplayAC = tempPower.ToString() + " Min.";
                            //MessageBox.Show(HibernateAC);

                        }

                        if (currentLine.Contains("DC"))
                        {
                            DisplayDC = currentLine;
                            DisplayDC = DisplayDC.Remove(0, 38);
                            tempPower = Int32.Parse(DisplayDC, System.Globalization.NumberStyles.HexNumber);
                            tempPower = tempPower / 60;
                            DisplayDC = tempPower.ToString() + " Min.";
                            //MessageBox.Show(HibernateDC);
                        }
                    }
                }

                File.Delete(@"c:\powerplanDisplay.txt");
                Display = DisplayDC + "," + DisplayAC;
            }

            if (SchoolDesktop)
            {
                ExecuteCommand(@"powercfg -query a8ec203a-6c54-4ebf-9b6e-64c621fcadf7 0012ee47-9041-4b5d-9b77-535fba8b1442 6738e2c4-e8a5-4a42-b16a-e040e769756e   > c:\powerplanHardDisk.txt");
                System.Threading.Thread.Sleep(2000);

                string HardAC = "";
                string HardDC = "";

                using (StreamReader reader = File.OpenText(@"c:\powerplanHardDisk.txt"))
                {
                    string currentLine;
                    while ((currentLine = reader.ReadLine()) != null)
                    {
                        if (currentLine.Contains("AC"))
                        {
                            HardAC = currentLine;
                            HardAC = HardAC.Remove(0, 38);
                            tempPower = Int32.Parse(HardAC, System.Globalization.NumberStyles.HexNumber);
                            tempPower = tempPower / 60;
                            HardAC = tempPower.ToString() + " Min.";
                            //MessageBox.Show(SleepAC);

                        }

                        if (currentLine.Contains("DC"))
                        {
                            HardDC = currentLine;
                            HardDC = HardDC.Remove(0, 38); // Parses and removes 0 -> 38 values from the currentLine string
                            tempPower = Int32.Parse(HardDC, System.Globalization.NumberStyles.HexNumber); // Converts from hexadecimal to int
                            tempPower = tempPower / 60; // Converts from seconds to minutes
                            HardDC = tempPower.ToString() + " Min.";
                            //MessageBox.Show(SleepDC);
                        }

                    }

                } // Power Plan Hard Disk Information

                File.Delete(@"c:\powerplanHardDisk.txt");
                HardDisk = HardDC + "," + HardAC;

                ExecuteCommand(@"powercfg -query a8ec203a-6c54-4ebf-9b6e-64c621fcadf7 238c9fa8-0aad-41ed-83f4-97be242c8f20 29f6c1db-86da-48c5-9fdb-f2b67b1f44da  >c:\powerplanSleep.txt");
                System.Threading.Thread.Sleep(2000);

                string SleepAC = "";
                string SleepDC = "";

                using (StreamReader reader = File.OpenText(@"c:\powerplanSleep.txt"))
                {
                    string currentLine;
                    while ((currentLine = reader.ReadLine()) != null)
                    {
                        if (currentLine.Contains("AC"))
                        {
                            SleepAC = currentLine;
                            SleepAC = SleepAC.Remove(0, 38);
                            tempPower = Int32.Parse(SleepAC, System.Globalization.NumberStyles.HexNumber);
                            tempPower = tempPower / 60;
                            SleepAC = tempPower.ToString() + " Min.";
                            //MessageBox.Show(SleepAC);

                        }

                        if (currentLine.Contains("DC"))
                        {
                            SleepDC = currentLine;
                            SleepDC = SleepDC.Remove(0, 38);
                            tempPower = Int32.Parse(SleepDC, System.Globalization.NumberStyles.HexNumber);
                            tempPower = tempPower / 60;
                            SleepDC = tempPower.ToString() + " Min.";
                            //MessageBox.Show(SleepDC);
                        }

                    }

                }

                File.Delete(@"c:\powerplanSleep.txt");
                SleepSettings = SleepDC + "," + SleepAC;

                ExecuteCommand(@"powercfg -query a8ec203a-6c54-4ebf-9b6e-64c621fcadf7 238c9fa8-0aad-41ed-83f4-97be242c8f20 9d7815a6-7ee4-497e-8888-515a05f02364 >c:\powerplanHibernate.txt");
                System.Threading.Thread.Sleep(2000);

                string HibernateAC = "";
                string HibernateDC = "";

                using (StreamReader reader = File.OpenText(@"c:\powerplanHibernate.txt"))
                {
                    string currentLine;
                    while ((currentLine = reader.ReadLine()) != null)
                    {
                        if (currentLine.Contains("AC"))
                        {
                            HibernateAC = currentLine;
                            HibernateAC = HibernateAC.Remove(0, 38);
                            tempPower = Int32.Parse(HibernateAC, System.Globalization.NumberStyles.HexNumber);
                            tempPower = tempPower / 60;
                            HibernateAC = tempPower.ToString() + " Min.";
                            //MessageBox.Show(HibernateAC);

                        }

                        if (currentLine.Contains("DC"))
                        {
                            HibernateDC = currentLine;
                            HibernateDC = HibernateDC.Remove(0, 38);
                            tempPower = Int32.Parse(HibernateDC, System.Globalization.NumberStyles.HexNumber);
                            tempPower = tempPower / 60;
                            HibernateDC = tempPower.ToString() + " Min.";
                            //MessageBox.Show(HibernateDC);
                        }
                    }
                }

                File.Delete(@"c:\powerplanHibernate.txt");
                HibernateSettings = HibernateDC + "," + HibernateAC;

                ExecuteCommand(@"powercfg -query a8ec203a-6c54-4ebf-9b6e-64c621fcadf7 4f971e89-eebd-4455-a8de-9e59040e7347 5ca83367-6e45-459f-a27b-476b1d01c936  >c:\powerplanLidClose.txt");
                System.Threading.Thread.Sleep(2000);
                string LidAC = "";
                string LidDC = "";

                using (StreamReader reader = File.OpenText(@"c:\powerplanLidClose.txt"))
                {
                    string currentLine;
                    while ((currentLine = reader.ReadLine()) != null)
                    {
                        if (currentLine.Contains("AC"))
                        {
                            LidAC = currentLine;
                            LidAC = LidAC.Remove(0, LidAC.Length - 3);
                            if (LidAC.Equals("000"))
                            {
                                LidAC = "Do nothing";
                            }
                            if (LidAC.Equals("001"))
                            {
                                LidAC = "Sleep";
                            }
                            if (LidAC.Equals("002"))
                            {
                                LidAC = "Hibernate";
                            }
                            if (LidAC.Equals("003"))
                            {
                                LidAC = "Shut down";
                            }
                        }

                        if (currentLine.Contains("DC"))
                        {
                            LidDC = currentLine;
                            LidDC = LidDC.Remove(0, LidDC.Length - 3);
                            if (LidDC.Equals("000"))
                            {
                                LidDC = "Do nothing";
                            }
                            if (LidDC.Equals("001"))
                            {
                                LidDC = "Sleep";
                            }
                            if (LidDC.Equals("002"))
                            {
                                LidDC = "Hibernate";
                            }
                            if (LidDC.Equals("003"))
                            {
                                LidDC = "Shut down";
                            }
                        }
                    }
                }

                File.Delete(@"c:\powerplanLidClose.txt");
                LidClose = LidDC + "," + LidAC;

                ExecuteCommand(@"powercfg -query a8ec203a-6c54-4ebf-9b6e-64c621fcadf7 4f971e89-eebd-4455-a8de-9e59040e7347 5ca83367-6e45-459f-a27b-476b1d01c936 >c:\powerplanButtonAction.txt");
                System.Threading.Thread.Sleep(2000);
                string PowerButtonAC = "";
                string PowerButtonDC = "";

                using (StreamReader reader = File.OpenText(@"c:\powerplanButtonAction.txt"))
                {
                    string currentLine;
                    while ((currentLine = reader.ReadLine()) != null)
                    {
                        if (currentLine.Contains("AC"))
                        {
                            PowerButtonAC = currentLine;
                            PowerButtonAC = PowerButtonAC.Remove(0, PowerButtonAC.Length - 3);
                            if (PowerButtonAC.Equals("000"))
                            {
                                PowerButtonAC = "Do nothing";
                            }
                            if (PowerButtonAC.Equals("001"))
                            {
                                PowerButtonAC = "Sleep";
                            }
                            if (PowerButtonAC.Equals("002"))
                            {
                                PowerButtonAC = "Hibernate";
                            }
                            if (PowerButtonAC.Equals("003"))
                            {
                                PowerButtonAC = "Shut down";
                            }
                            if (PowerButtonAC.Equals("004"))
                            {
                                PowerButtonAC = "Turn off display";
                            }

                        }

                        if (currentLine.Contains("DC"))
                        {
                            PowerButtonDC = currentLine;
                            PowerButtonDC = PowerButtonDC.Remove(0, PowerButtonDC.Length - 3);
                            if (PowerButtonDC.Equals("000"))
                            {
                                PowerButtonDC = "Do nothing";
                            }
                            if (PowerButtonDC.Equals("001"))
                            {
                                PowerButtonDC = "Sleep";
                            }
                            if (PowerButtonDC.Equals("002"))
                            {
                                PowerButtonDC = "Hibernate";
                            }
                            if (PowerButtonDC.Equals("003"))
                            {
                                PowerButtonDC = "Shut down";
                            }
                            if (PowerButtonDC.Equals("004"))
                            {
                                PowerButtonDC = "Turn off display";
                            }

                        }
                    }
                }

                File.Delete(@"c:\powerplanButtonAction.txt");
                PowerButton = PowerButtonDC + "," + PowerButtonAC;

                ExecuteCommand(@"powercfg -query a8ec203a-6c54-4ebf-9b6e-64c621fcadf7 7516b95f-f776-4464-8c53-06167f40cc99 3c0bc021-c8a8-4e07-a973-6b14cbcb2b7e  >c:\powerplanDisplay.txt");
                System.Threading.Thread.Sleep(2000);

                string DisplayAC = "";
                string DisplayDC = "";

                using (StreamReader reader = File.OpenText(@"c:\powerplanDisplay.txt"))
                {
                    string currentLine;
                    while ((currentLine = reader.ReadLine()) != null)
                    {
                        if (currentLine.Contains("AC"))
                        {
                            DisplayAC = currentLine;
                            DisplayAC = DisplayAC.Remove(0, 38);
                            tempPower = Int32.Parse(DisplayAC, System.Globalization.NumberStyles.HexNumber);
                            tempPower = tempPower / 60;
                            DisplayAC = tempPower.ToString() + " Min.";
                            //MessageBox.Show(HibernateAC);

                        }

                        if (currentLine.Contains("DC"))
                        {
                            DisplayDC = currentLine;
                            DisplayDC = DisplayDC.Remove(0, 38);
                            tempPower = Int32.Parse(DisplayDC, System.Globalization.NumberStyles.HexNumber);
                            tempPower = tempPower / 60;
                            DisplayDC = tempPower.ToString() + " Min.";
                            //MessageBox.Show(HibernateDC);
                        }
                    }
                }

                File.Delete(@"c:\powerplanDisplay.txt");
                Display = DisplayDC + "," + DisplayAC;
            }

            //System properties verification Starts
            string Run = "";
            DialogResult SystemDialog = MessageBox.Show(new Form { TopMost = true }, "Is run included in the start menu? Press the Windows key or click on the Windows icon at the bottom left of the screen.", "Run Verification", MessageBoxButtons.YesNoCancel);
            if (SystemDialog == DialogResult.Yes)
            {
                Run = "YES";
            }
            if (SystemDialog == DialogResult.No)
            {
                Run = "NO";
            }
            if (SystemDialog == DialogResult.Cancel)
            {
                Run = "Not Applicable";
            }

            string EnterpriseMode = "";
            RegistryKey Enterprise = Registry.LocalMachine.OpenSubKey(@"SOFTWARE\Policies\Microsoft\Internet Explorer\Main\EnterpriseMode");
            if (Enterprise != null)
            {
                EnterpriseMode = "YES is available";
                Enterprise.Close();
            }
            if (Enterprise == null)
            {
                EnterpriseMode = "NO is not available";
            }

            string HomePage = "";

            try
            {

                Process.Start(@"C:\Program Files (x86)\Internet Explorer\iexplore.exe");
                System.Threading.Thread.Sleep(2000);
                SystemDialog = MessageBox.Show(new Form { TopMost = true }, "Is home page set to http://schools.nyc.gov?", "Home Page Verification", MessageBoxButtons.YesNoCancel);
                if (SystemDialog == DialogResult.Yes)
                {
                    HomePage = "YES";
                }
                if (SystemDialog == DialogResult.No)
                {
                    HomePage = "NO";
                }
                if (SystemDialog == DialogResult.Cancel)
                {
                    HomePage = "Not Applicable";
                }

            }
            
            catch
            {
                HomePage = "Internet Explorer .exe file could not be found.";
            }

            string Favorites = "";
            try
            {

                Process.Start(@"C:\Users\nycdoe\Favorites");
                //string Favorites = "";
                SystemDialog = MessageBox.Show(new Form { TopMost = true }, "Verify these favorites:" + "\r\n" +
                    "http://ats.nycboe.net/atsprint/" + "\r\n" +
                    "http://go.microsoft.com/fwlink/p/?LinkId=255142" + "\r\n" + "https://nycdoe.cybershift.net/"
                    + "\r\n" + "http://mygalaxy.nycenet.edu/" + "\r\n" + "https://servicecenter.nycenet.edu/selfsupport/"
                    + "\r\n" + "http://intranet.nycboe.net/sandbox/default.htm" + "\r\n" + "https://wc.nycenet.edu/", "Favorites Verification", MessageBoxButtons.YesNoCancel);
                if (SystemDialog == DialogResult.Yes)
                {
                    Favorites = "YES";
                }
                if (SystemDialog == DialogResult.No)
                {
                    Favorites = "NO";
                }
                if (SystemDialog == DialogResult.Cancel)
                {
                    Favorites = "Not Applicable";
                }
            }
            catch
            {
                Favorites = "NYCDOE Favorites Folder could not be found";
            }

            string ProxyIE = "";
            try
            {

                ExecuteCommand("inetcpl.cpl,4");
                SystemDialog = MessageBox.Show(new Form { TopMost = true }, "Verify NO PROXY set in Internet options connections tab", "Proxy Verification", MessageBoxButtons.YesNoCancel);
                if (SystemDialog == DialogResult.Yes)
                {
                    ProxyIE = "Yes there is no proxy set";
                }
                if (SystemDialog == DialogResult.No)
                {
                    ProxyIE = "No there is a proxy set";
                }
                if (SystemDialog == DialogResult.Cancel)
                {
                    ProxyIE = "Not Applicable";
                }
            }
            catch
            {
                ProxyIE = "ProxyIE command: inetcpl.cpl,4 , could not be started";
            }

            string PopUpBlocker = "";
            try
            {
                ExecuteCommand("inetcpl.cpl,2");
                SystemDialog = MessageBox.Show(new Form { TopMost = true }, "Verify pop-up blocker settings (GO TO SETTINGS IN PRIVACY TAB) to medium and that the following allowed sites are added:" +
                    "\r\n" + ".carnegielearning.com" + "\r\n" + ".cybershift.net" + "\r\n" + ".learning.aries.net"
                    + "\r\n" + ".nycboe.net" + "\r\n" + ".nycenet.edu" + "\r\n" + ".platform.ilearnnyc.net" + "\r\n" + ".playdigit.org"
                    + "\r\n" + ".thelearningodyssey.com", "Proxy Verification", MessageBoxButtons.YesNoCancel);
                if (SystemDialog == DialogResult.Yes)
                {
                    PopUpBlocker = "Yes pop-up blocker set to medium and the listed sites are allowed";
                }
                if (SystemDialog == DialogResult.No)
                {
                    PopUpBlocker = "No pop-up blocker is not set to medium and the listed sites are allowed";
                }
                if (SystemDialog == DialogResult.Cancel)
                {
                    PopUpBlocker = "Not Applicable";
                }
            }
            catch
            {
                PopUpBlocker = "PopUpBlockerIE command: inetcpl.cpl,2 , could not be started";
            }

            MessageBox.Show("This message will self destruct when you press okay and will continue with verification, stand by....");

            string NYCDOE = "";
            try
            {

                //User Account Verification
                ExecuteCommand(@"NET USER nycdoe >c:\nycdoe.txt");
                System.Threading.Thread.Sleep(10000);
                using (StreamReader reader = File.OpenText(@"c:\nycdoe.txt"))
                {
                    string currentLine;
                    while ((currentLine = reader.ReadLine()) != null)
                    {
                        if (currentLine.Contains("Password expires"))
                        {
                            NYCDOE = currentLine;

                        }
                    }
                }

                File.Delete(@"c:\nycdoe.txt");

            }
            catch
            {
                NYCDOE = "NYCDOE Account could not be found.";
            }

            string teacher = "";
            try
            {
                ExecuteCommand(@"NET USER teacher >c:\teacher.txt");
                System.Threading.Thread.Sleep(10000);
                using (StreamReader reader = File.OpenText(@"c:\teacher.txt"))
                {
                    string currentLine;
                    while ((currentLine = reader.ReadLine()) != null)
                    {
                        if (currentLine.Contains("Password expires"))
                        {
                            teacher = currentLine;

                        }
                    }
                }

                File.Delete(@"c:\teacher.txt");
            }
            catch
            {
                teacher = "Teacher Account could not be found.";
            }

            string student = "";
            try
            {

                ExecuteCommand(@"NET USER student >c:\student.txt");
                System.Threading.Thread.Sleep(10000);
                string temp1 = "";
                string temp2 = "";
                using (StreamReader reader = File.OpenText(@"c:\student.txt"))
                {
                    string currentLine;
                    while ((currentLine = reader.ReadLine()) != null)
                    {
                        if (currentLine.Contains("Password expires"))
                        {
                            temp1 = currentLine;

                        }

                        if (currentLine.Contains("User may"))
                        {
                            temp2 = currentLine;

                        }
                    }
                }

                File.Delete(@"c:\student.txt");
                student = temp1 + "," + temp2;
            }
            catch
            {
                student = "Student Account could not be found";
            }


            StringBuilder WirelessdriverInfo = new StringBuilder();
            StringBuilder DisplaydriverInfo = new StringBuilder();

            ManagementObjectSearcher searcher = new ManagementObjectSearcher("root\\CIMV2", "SELECT * FROM Win32_PnPSignedDriver");
            ManagementObjectCollection moc = searcher.Get();

            string Name = "";
            string Date = "";
            DateTime Driverdate;
            try
            {
                foreach (var manObj in moc)
                {
                    if (manObj["FriendlyName"] != null)
                    {
                        Name = manObj["FriendlyName"].ToString();

                        if ((Name.Contains("Intel(R) Dual Band")))
                        {
                            Date = manObj["DriverDate"].ToString();
                            Date = Date.Substring(0, 8);
                            Driverdate = DateTime.ParseExact(Date, "yyyyMMdd", null);
                            Date = Driverdate.ToString().Substring(0, 10);
                            WirelessdriverInfo.Append("Device Name: " + manObj["FriendlyName"] + "," + "DriverVersion: " + manObj["DriverVersion"] + "," + "DriverDate: " + Date);
                        } // Wireless driver
                    }
                }
            }
            catch (ManagementException exception)
            {
                MessageBox.Show("An error occurred while querying for WMI Wireless driver data: " + exception.Message);
            }

            try
            {
                searcher =
                    new ManagementObjectSearcher("root\\CIMV2",
                    "SELECT * FROM Win32_VideoController");

                foreach (ManagementObject manObj in searcher.Get())
                {
                    Date = manObj["DriverDate"].ToString();
                    Date = Date.Substring(0, 8);
                    Driverdate = DateTime.ParseExact(Date, "yyyyMMdd", null);
                    Date = Driverdate.ToString().Substring(0, 10);
                    DisplaydriverInfo.AppendLine("Display driver information: ," + "Device Name: " + manObj["Name"] + "," + "DriverVersion: " + manObj["DriverVersion"] + "," + "DriverDate: " + Date);
                    //MessageBox.Show(DisplaydriverInfo.ToString());
                }
            }
            catch (ManagementException exception)
            {
                MessageBox.Show("An error occurred while querying for WMI Video display data: " + exception.Message);
            }

            /**ExecuteCommand(@"wmic PATH Win32_videoController GET description, driverversion, driverdate > C:\video.txt");
            System.Threading.Thread.Sleep(1000);

            string videoDriver = "";
            string videoDate = "";
            string videoNumber = "";

            using (StreamReader reader = File.OpenText(@"c:\video.txt"))
            {
                string currentLine;
                while ((currentLine = reader.ReadLine()) != null)
                {
                    if (currentLine.Contains("Intel"))
                    {
                        videoDriver = currentLine;
                        videoDriver = videoDriver.Remove(34, 17);
                        videoDate = videoDriver.Substring(26, 8);
                        videoNumber = videoDriver.Substring(36, 13);
                        videoDriver = videoDriver.Substring(0, 25);
                    }
                }
            }

            File.Delete(@"c:\video.txt");

            try
            {
                Date = videoDate;
                Date = Date.Substring(0, 8);
                Date = Date.Substring(0, 8);
                Driverdate = DateTime.ParseExact(Date, "yyyyMMdd", null);
                videoDate = Driverdate.ToString();
                videoDate = videoDate.Substring(0, 10);

                //videoDriver = "Device Name: " + videoDriver + "," + "Driver Version: " + videoNumber + "," + "Driver Date: " + videoDate;
            }

            catch
            {
                Date = videoDate;
            }

            videoDriver = "Device Name: " + videoDriver + "," + "Driver Version: " + videoNumber + "," + "Driver Date: " + videoDate; **/

            ExecuteCommand("sysdm.cpl");

            string RemoteDesktop = "";
            string SysProtection = "";
            string AutoRestart = "";
            string TaskBar = "";
            string Autoplay = "";

            System.Threading.Thread.Sleep(3000);
            DialogResult dialogResult = MessageBox.Show(new Form { TopMost = true }, "Is Remote Desktop access enabled?", "Remote Access Verification", MessageBoxButtons.YesNoCancel);
            if (dialogResult == DialogResult.Yes)
            {
                RemoteDesktop = "YES";
            }
            else if (dialogResult == DialogResult.No)
            {
                RemoteDesktop = "NO";
            }
            else if (dialogResult == DialogResult.Cancel)
            {
                RemoteDesktop = "Not Applicable";
            }

            dialogResult = MessageBox.Show(new Form { TopMost = true }, "Is System Protection disabled?", "System Protection Verification", MessageBoxButtons.YesNoCancel);
            if (dialogResult == DialogResult.Yes)
            {
                SysProtection = "YES";
            }
            else if (dialogResult == DialogResult.No)
            {
                SysProtection = "NO";
            }
            else if (dialogResult == DialogResult.Cancel)
            {
                SysProtection = "Not Applicable";
            }

            dialogResult = MessageBox.Show(new Form { TopMost = true }, "Is automatically restart unchecked?", "Automatically restart Verification", MessageBoxButtons.YesNoCancel);
            if (dialogResult == DialogResult.Yes)
            {
                AutoRestart = "YES";
            }
            else if (dialogResult == DialogResult.No)
            {
                AutoRestart = "NO";
            }
            else if (dialogResult == DialogResult.Cancel)
            {
                AutoRestart = "Not Applicable";
            }

            // add more N/A buttons after this didnt FINISH!!!

            ExecuteCommand("start ms-settings:taskbar");
            dialogResult = MessageBox.Show(new Form { TopMost = true }, "Are taskbar buttons set to never combine?", "Taskbar Verification", MessageBoxButtons.YesNoCancel);
            if (dialogResult == DialogResult.Yes)
            {
                TaskBar = "YES";
            }
            else if (dialogResult == DialogResult.No)
            {
                TaskBar = "NO";
            }
            else if (dialogResult == DialogResult.Cancel)
            {
                TaskBar = "N/A";
            }


            ExecuteCommand(@"control /name Microsoft.Autoplay");
            dialogResult = MessageBox.Show(new Form { TopMost = true }, "Is autoplay unchecked?", "Autoplay Verification", MessageBoxButtons.YesNoCancel);
            if (dialogResult == DialogResult.Yes)
            {
                Autoplay = "YES";
            }
            else if (dialogResult == DialogResult.No)
            {
                Autoplay = "NO";
            }
            else if (dialogResult == DialogResult.Cancel)
            {
                Autoplay = "N/A";
            }


            //Registry Verication Begins
            RegistryKey RK = Registry.LocalMachine.OpenSubKey(@"SYSTEM\CurrentControlSet\Control\Lsa");
            string RestrictAnonymous = "";
            if (RK != null)
            {
                RestrictAnonymous = "Register value is: " + RK.GetValue("restrictanonymous").ToString();
                RK.Close();
            }
            if (RK == null)
            {
                RestrictAnonymous = "Register value is null";
            }

            RK = Registry.LocalMachine.OpenSubKey(@"SOFTWARE\Intel\Wireless\AT6");
            string AMT = "";
            if (RK != null)
            {
                AMT = "Register value is: " + RK.GetValue("iAMTe").ToString();
                RK.Close();
            }
            if (RK == null)
            {
                AMT = "Register value is null";
            }

            RK = Registry.LocalMachine.OpenSubKey(@"SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon\SpecialAccounts\UserList");
            string users = "";

            if (RK != null)
            {
                users = "Register DOES exist";
                RK.Close();
            }

            if (RK == null)
            {
                users = "Register DOES NOT exist";
            }


            RK = Registry.LocalMachine.OpenSubKey(@"SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System");
            string AdminShares = "";

            if (RK != null)
            {
                if (RK.GetValue("LocalAccountTokenFilterPolicy") == null)
                {
                    AdminShares = "LocalAccountTokenFilterPolicy does not exist";
                }

                if (RK.GetValue("LocalAccountTokenFilterPolicy") != null)
                {
                    if (RK.GetValue("LocalAccountTokenFilterPolicy").ToString().Equals("1"))
                    {
                        AdminShares = "Admin Shares is ENABLED with value: " + RK.GetValue("LocalAccountTokenFilterPolicy").ToString();
                    }

                    if (RK.GetValue("LocalAccountTokenFilterPolicy").ToString().Equals("0"))
                    {
                        AdminShares = "Admin Shares is DISABLED with value: " + RK.GetValue("LocalAccountTokenFilterPolicy").ToString();
                    }
                }

                RK.Close();
            }

            if (RK == null)
            {
                AdminShares = "Register does not exist";
            }


            //Security Verification
            string LockOutPolicy = "";
            string DeviceDrivers = "";
            string Firewall = "";
            string SCCM = "";
            string Proxy = "";
            string Cortana = "";

            try
            {

                Process.Start("gpedit.msc");
                System.Threading.Thread.Sleep(2000);

                dialogResult = MessageBox.Show(new Form { TopMost = true }, "Verify password account lockout policy. Go to Computer Configuration -> Windows settings -> Security settings -> Account Policies -> Account Lockout Policy.",
                        "Lockout Policy Verification", MessageBoxButtons.YesNoCancel);
                if (dialogResult == DialogResult.Yes)
                {
                    LockOutPolicy = "YES";
                }
                else if (dialogResult == DialogResult.No)
                {
                    LockOutPolicy = "NO";
                }
                else if (dialogResult == DialogResult.Cancel)
                {
                    LockOutPolicy = "N/A";
                }

                dialogResult = MessageBox.Show(new Form { TopMost = true }, "Verify Power Users can load and unload device drivers. Go to Computer Configuration -> Windows settings -> Security settings -> Local Policies -> User Rights Assignment -> load and unload device drivers.",
                        "Power User Policy Verification", MessageBoxButtons.YesNoCancel);
                if (dialogResult == DialogResult.Yes)
                {
                    DeviceDrivers = "YES";
                }
                else if (dialogResult == DialogResult.No)
                {
                    DeviceDrivers = "NO";
                }
                else if (dialogResult == DialogResult.No)
                {
                    DeviceDrivers = "N/A";
                }

            }

            catch
            {
                MessageBox.Show("Could not display gpedit.");
            }

            System.Threading.Thread.Sleep(2000);

            try
            {
                RK = Registry.LocalMachine.OpenSubKey(@"Software\policies\Microsoft\WindowsFirewall\StandardProfile");
                if (RK == null)
                {
                    Firewall = "Could not find register. Please try gpedit.msc in the run application";
                }

                if (RK != null)
                {
                    if (RK.GetValue("EnableFirewall").ToString().Equals("1"))
                    {
                        Firewall = "Enabled with value 1";
                    }

                    else if (RK.GetValue("EnableFirewall").ToString().Equals("0"))
                    {
                        Firewall = "Disabled with value 0";
                    }

                    RK.Close();
                }
            }
            catch
            {
                Firewall = "Could not find register. Please try gpedit.msc in the run application";
            }

            try
            {

                RK = Registry.LocalMachine.OpenSubKey(@"SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate");
                if (RK == null)
                {
                    SCCM = "Could not find register. Please try gpedit.msc in the run application";
                }

                if (RK != null)
                {
                    if (RK.GetValue("AcceptTrustedPublisherCerts").ToString().Equals("1"))
                    {
                        SCCM = "Enabled with value 1";
                    }

                    else if (RK.GetValue("AcceptTrustedPublisherCerts").ToString().Equals("0"))
                    {
                        SCCM = "Disabled with value 0";
                    }

                    RK.Close();

                }
            }
            catch
            {
                SCCM = "Could not find register. Please try gpedit.msc in the run application";
            }


            try
            {
                RK = Registry.LocalMachine.OpenSubKey(@"SOFTWARE\Policies\Microsoft\Windows\Windows Search");
                if (RK == null)
                {
                    Cortana = "Could not find register. Please try gpedit.msc in the run application";
                }

                if (RK != null)
                {
                    if (RK.GetValue("AllowCortana").ToString().Equals("1"))
                    {
                        Cortana = "Enabled with value 1";
                    }

                    else if (RK.GetValue("AllowCortana").ToString().Equals("0"))
                    {
                        Cortana = "Disabled with value 0";
                    }

                    RK.Close();

                }
            }
            catch
            {
                Cortana = "Could not find register. Please try gpedit.msc in the run application";
            }

            try
            {
                ExecuteCommand("netsh winhttp show proxy >c:/proxy.txt");

                System.Threading.Thread.Sleep(2000);

                using (StreamReader reader = File.OpenText(@"c:\proxy.txt"))
                {
                    string currentLine;
                    while ((currentLine = reader.ReadLine()) != null)
                    {
                        if (currentLine.Contains("Direct"))
                        {
                            Proxy = currentLine;
                        }

                    }

                }

                File.Delete(@"c:\proxy.txt");
                Proxy = Proxy.TrimStart();
            }

            catch
            {
                Proxy = "Proxy command could not be run. Try to run this command: netsh winhttp show proxy in the cmd and check the result";
            }

            string Activation = "";

            dialogResult = MessageBox.Show(new Form { TopMost = true }, "Did Windows and Office activate? If not make sure computer is connected to the network.",
                        "Windows and Office activation Verification", MessageBoxButtons.YesNoCancel);
            if (dialogResult == DialogResult.Yes)
            {
                Activation = "YES";
            }
            else if (dialogResult == DialogResult.No)
            {
                Activation = "NO";

            }
            else if (dialogResult == DialogResult.Cancel)
            {
                Activation = "Not Applicable";
            }

            string PrintingCheck = "";
            try
            {
                Process.Start(@"C:\Program Files (x86)\Microsoft Office\Office16\WINWORD.EXE");
                System.Threading.Thread.Sleep(1000);
                dialogResult = MessageBox.Show(new Form { TopMost = true }, "Can you print using Microsoft Word?",
                            "Microsoft Word Printing Verification", MessageBoxButtons.YesNoCancel);
                if (dialogResult == DialogResult.Yes)
                {
                    PrintingCheck = "YES";
                }
                else if (dialogResult == DialogResult.No)
                {
                    PrintingCheck = "NO";
                }
                else if (dialogResult == DialogResult.Cancel)
                {
                    PrintingCheck = "Not Applicable";
                }
            }
            catch
            {
                PrintingCheck = "Could not find Microsoft Word";
            }

            string ConfigurationCheck = "";
            try
            {
                ExecuteCommand("control smscfgrc");
                dialogResult = MessageBox.Show(new Form { TopMost = true }, "Did Configuration Manager Properties(SCCM) open?",
                            "Configuration Manager Verification", MessageBoxButtons.YesNoCancel);
                if (dialogResult == DialogResult.Yes)
                {
                    ConfigurationCheck = "YES";
                }
                else if (dialogResult == DialogResult.No)
                {
                    ConfigurationCheck = "NO";
                }
                else if (dialogResult == DialogResult.Cancel)
                {
                    ConfigurationCheck = "Not Applicable";
                }
            }
            catch
            {
                ConfigurationCheck = "Could not find SCCM";
            }

            string SymantecCheck = "";
            string SymantecPath = "";
            try
            {
                ExecuteCommand(@"wmic process where ""name = 'ccSvcHst.exe'"" get ExecutablePath >C:\symantec.txt");
                System.Threading.Thread.Sleep(5000);
                using (StreamReader reader = File.OpenText(@"C:\symantec.txt"))
                {
                    string currentLine;
                    while ((currentLine = reader.ReadLine()) != null)
                    {
                        if (currentLine.Contains(@"C:\"))
                        {
                            SymantecPath = currentLine;
                        }
                    }
                }

                File.Delete(@"C:\symantec.txt");

                SymantecPath = SymantecPath.Remove(83, 15);
                SymantecPath = SymantecPath + @"\SymCorpUI.exe";

                //Process.Start(@"C:\Program Files (x86)\Symantec\Symantec Endpoint Protection\14.0.3897.1101.105\Bin\SymCorpUI.exe");
                Process.Start(SymantecPath);
                dialogResult = MessageBox.Show(new Form { TopMost = true }, "Please load up Symantec and check that all Symantec Protection Modules are green and checked. You may need to wait a few moments (it should update as long as the machine is on the network).",
                            "Symantec Protection Modules Verification", MessageBoxButtons.YesNoCancel);
                if (dialogResult == DialogResult.Yes)
                {
                    SymantecCheck = "YES";
                }
                else if (dialogResult == DialogResult.No)
                {
                    SymantecCheck = "NO";
                }
                else if (dialogResult == DialogResult.Cancel)
                {
                    SymantecCheck = "Not Applicable";
                }
            }
            catch
            {
                SymantecCheck = "Could not find Symantec";
            }


            string SymantecResetCheck = "";

            try
            {
                Process.Start(@"C:\ProgramData\Symantec\Symantec Endpoint Protection\PersistedData");
                System.Threading.Thread.Sleep(1000);
                dialogResult = MessageBox.Show(new Form { TopMost = true }, "Verify Symantec has reset the hardware ID using sephwid.xml, should have been created when machine was imaged.",
                            "Symantec Reset Verification", MessageBoxButtons.YesNoCancel);
                if (dialogResult == DialogResult.Yes)
                {
                    SymantecResetCheck = "YES";
                }
                else if (dialogResult == DialogResult.No)
                {
                    SymantecResetCheck = "NO";
                }
                else if (dialogResult == DialogResult.Cancel)
                {
                    SymantecResetCheck = "Not Applicable";
                }
            }
            catch
            {
                SymantecResetCheck = "Could not find Symantec";
            }

            string SymantecServerCheck = "";

            try
            {
                Process.Start(@"C:\ProgramData\Symantec\Symantec Endpoint Protection\CurrentVersion\Data\Config\SyLink.xml");
                System.Threading.Thread.Sleep(1000);
                dialogResult = MessageBox.Show(new Form { TopMost = true }, "Verify Symantec has the correct server settings",
                            "Symantec Server Settings Verification", MessageBoxButtons.YesNoCancel);
                if (dialogResult == DialogResult.Yes)
                {
                    SymantecServerCheck = "YES";
                }
                else if (dialogResult == DialogResult.No)
                {
                    SymantecServerCheck = "NO";
                }
                else if (dialogResult == DialogResult.Cancel)
                {
                    SymantecServerCheck = "Not Applicable";
                }
            }
            catch
            {
                SymantecServerCheck = "Could not find Symantec";
            }

            string HotkeysCheck = "";

            try
            {
                dialogResult = MessageBox.Show(new Form { TopMost = true }, "Verify Hotkeys and FN keys work properly.",
                            "Hotkeys and FN keys Verification", MessageBoxButtons.YesNoCancel);
                if (dialogResult == DialogResult.Yes)
                {
                    HotkeysCheck = "YES";
                }
                else if (dialogResult == DialogResult.No)
                {
                    HotkeysCheck = "NO";
                }
                else if (dialogResult == DialogResult.Cancel)
                {
                    HotkeysCheck = "Not Applicable";
                }
            }
            catch
            {
                HotkeysCheck = "Could not find Symantec";
            }

            string PowerDVD = "";
            try
            {
                Process.Start(@"C:\Program Files (x86)\CyberLink");
                System.Threading.Thread.Sleep(1000);
                dialogResult = MessageBox.Show(new Form { TopMost = true }, "Verify POWERDVD is in the image.",
                            "POWERDVD Verification", MessageBoxButtons.YesNoCancel);
                if (dialogResult == DialogResult.Yes)
                {
                    PowerDVD = "YES";

                }
                else if (dialogResult == DialogResult.No)
                {
                    PowerDVD = "NO";
                }
                else if (dialogResult == DialogResult.Cancel)
                {
                    PowerDVD = "Not Applicable";
                }
            }
            catch
            {
                PowerDVD = "Could not find POWERDVD file path";
            }

            string UACCheck = "";

            try
            {
                Process.Start(@"C:\Windows\System32\UserAccountControlSettings.exe");
                System.Threading.Thread.Sleep(1000);
                dialogResult = MessageBox.Show(new Form { TopMost = true }, "Verify UAC is turned on and set to the recommended setting (third from the bottom).",
                            "UAC Verification", MessageBoxButtons.YesNoCancel);
                if (dialogResult == DialogResult.Yes)
                {
                    UACCheck = "YES";
                    
                }
                else if (dialogResult == DialogResult.No)
                {
                    UACCheck = "NO";
                }
                else if (dialogResult == DialogResult.Cancel)
                {
                    UACCheck = "Not Applicable";
                }
            }
            catch
            {
                UACCheck = "Could not find UAC command";
            }

            MessageBox.Show("System will now reset");

            Process[] AllProcesses = Process.GetProcesses();

            foreach (var p in AllProcesses)
            {
                string Process = p.ProcessName.ToUpper();
                if (Process == "IEXPLORE")
                {
                    p.Kill();
                }
                if (Process == "MICROSOFTEDGE")
                {
                    p.Kill();
                }
                if (Process == "SystemPropertiesComputerName".ToUpper())
                {
                    p.Kill();
                }
                if (Process == "SystemSettings".ToUpper())
                {
                    p.Kill();
                }
                if (Process == "WINWORD".ToUpper())
                {
                    p.Kill();
                }
                if (Process == "mmc".ToUpper())
                {
                    p.Kill();
                }
                if (Process == "dllhost".ToUpper())
                {
                    p.Kill();
                }
                if (Process == "rundll32".ToUpper())
                {
                    p.Kill();
                }
                if (Process == "SymCorpUI".ToUpper())
                {
                    p.Kill();
                }
                if (Process == "explorer".ToUpper())
                {
                    p.Kill();
                }
            }

            Process.Start("explorer.exe");

            String[] Header = {
               "PRELIMINARY VERIFICATION HAS STARTED","Computer Name (Should be first initial of vendor name + serial):", "BIOS Information:", "Windows License Information (Verify Key is BY6DY for school and central laptops):",
                "Office Information: ", "Verify DOE Installs Folder in C drive:","Verify version 1.3 or newer", "Verify image version matches image information file ", "",

                "SYSTEM PROPERTIES VERIFICATION HAS STARTED", DisplaydriverInfo.ToString(), "Wireless driver information", "Is Remote Desktop on: ",
                "Is System Protection turned off: ", "Is automatically restart unchecked: ", "Are taskbar buttons set to never combine?", "Verify Autoplay is unchecked: ", "",

                "POWER VERFICATION HAS STARTED", "Settings", "Hard Disk -> Turn Off Hard Disk After","Sleep -> Sleep After", "Hibernate After", "Power Button & Lid -> Lid Close Action",
                "Power Button & Lid -> Power Button Action", "Display -> Turn Off Display After", "",

                "SYSTEM PROPERTIES VERIFICATION CONTINUED","Verify Run is included in the Start Menu: ","Verify Enterprise Mode: ",
                "Verify Home Page is set to: http://schools.nyc.gov", "Verify NO PROXY set in the Internet options connections tab",
                "Verify pop-up blocker settings to medium and that the listed allowed sites are added:","Verify Favorites:",

                "",

                "USER VERFICATION HAS STARTED","NYCDOE: ", "Teacher: ", "Student: ",

                "","REGISTRY VERIFICATION HAS STARTED", "RestrictAnonymous register value is:",
                "AMT register value is:", "Does the UserList register path exist?", "Verify Admin Shares is enabled:",

                "","SECURITY VERIFICATION HAS STARTED", "Verify Lockout Policy: ", "Verify Load and Unload Device Drivers: ", "Verify Firewall is disabled (meaning it is enabled): ",
                "Verify system is configured to work with SCCM", "Verify there is no PROXY(should say direct)", "Verify that Cortana is disabled",

                "",
                "TO ENSURE COMPLETE TESTING:",
                "Windows and Office must be able to activate (must be on POC network connected)",
                "Verify printing can occur from Microsoft Word; physically printing is not required",
                "Verify that all Symantec protection modules are green and checked.",
                "When logged into the Student account verify DOE wallpaper cannot be changed. Please check the Image CheckList sheet for other requirements.",
                "When logged into the Student account verify that the Start Menu cannot be changed",
                "Verify Configuration Manager(SCCM) is located in the Control Panel",
                "Verify Symantec has reset the hardware ID",
                "Verify Symantec has the correct server settings",
                "Confirm Hotkeys and FN keys work properly",
                "Verify Admin Shares is Enabled with a value of 1",
                "Ensure Cortana is disabled",
                "Confirm that UAC is turned on the Recommended setting (third from the bottom)",
                "Confirm POWERDVD is installed",

            };

            String[] Data = {
               "", Environment.MachineName, BiosData, MAKData, OfficeData, DOEinstalls, DOEinstallsVersion, Image_Information,"", "", "", WirelessdriverInfo.ToString(),
                RemoteDesktop,SysProtection,AutoRestart, TaskBar,Autoplay,"","","Running On Batteries, Plugged In",
                HardDisk, SleepSettings, HibernateSettings, LidClose, PowerButton, Display, "",
                "",Run,EnterpriseMode,HomePage,ProxyIE,PopUpBlocker,Favorites,"",
                "", NYCDOE, teacher, student,
                "", "",RestrictAnonymous, AMT, users,
                AdminShares, "", "", LockOutPolicy, DeviceDrivers, Firewall, SCCM, Proxy, Cortana,

                "",
                "",
                Activation,
                PrintingCheck,
                SymantecCheck,
                "Cannot verify must be logged into the student account",
                "Cannot verify must be logged into the student account",
                ConfigurationCheck,
                SymantecResetCheck,
                SymantecServerCheck,
                HotkeysCheck,
                "ALREADY VERIFIED ABOVE IN REGISTRY VERIFICATION",
                "ALREADY VERIFIED ABOVE IN SECURITY SETTINGS",
                UACCheck,
                PowerDVD,

            };

            if (!File.Exists(@"C:\AppVerify.csv"))
            {
                using (StreamWriter stream = File.CreateText(@"C:\AppVerify.csv")) {

                    DateTime CurrentDay = DateTime.Today;
                    string Trademark = string.Format("{0},{1}", "AppVerify made by SD&E Devices and Imaging", "Current Date: " + CurrentDay.ToString("d"));
                    RK = Registry.LocalMachine.OpenSubKey(@"HARDWARE\DESCRIPTION\System\BIOS");
                    string SystemHardware = RK.GetValue("SystemFamily").ToString();
                    stream.WriteLine(Trademark);
                    stream.WriteLine(SystemHardware);
                    stream.WriteLine(IP);
                    stream.WriteLine();

                    for (int i = 0; i < Header.Count(); i++)
                    {
                        string first = Header[i].ToString();
                        string second = Data[i];
                        string csvRow = string.Format("{0},{1}", first, second);

                        stream.WriteLine(csvRow);
                    }

                    try
                    {

                        stream.WriteLine();
                        stream.WriteLine("------- Installed Applications -------");

                        int MissingProcessCount = 0;
                        pattern = "\\s+";
                        string replacement = " ";
                        Regex rgx = new Regex(pattern);

                        string CommandInstalledApps = @"wmic product get name, version >C:\allprograms.txt";
                        ExecuteProgramsCommand(CommandInstalledApps);
                        System.Threading.Thread.Sleep(5000);
                        String[] InstalledProgramsList = new String[27];

                        for (int i = 0; i <= 26; i++)
                        {
                            InstalledProgramsList[i] = "";
                        }

                        InstalledProgramsList[24] = "------- Program Features Not Listed -------";

                        System.Threading.Thread.Sleep(50000);

                        using (StreamReader reader = File.OpenText(@"c:\allprograms.txt"))
                        {
                            string currentLine;
                            while ((currentLine = reader.ReadLine()) != null)
                            {

                                if (currentLine.Contains("Adobe AIR"))
                                {
                                    currentLine = currentLine.Insert(9, ",");
                                    InstalledProgramsList[0] = currentLine;
                                    InstalledProgramsList[0] = rgx.Replace(InstalledProgramsList[0], replacement);
                                }

                                if (currentLine.Contains("Adobe Acrobat"))
                                {
                                    currentLine = currentLine.Insert(23, ",");
                                    InstalledProgramsList[1] = currentLine;
                                    InstalledProgramsList[1] = rgx.Replace(InstalledProgramsList[1], replacement);
                                }

                                if (currentLine.Contains("Adobe Shockwave"))
                                {
                                    //currentLine = currentLine.Insert(23, ",");
                                    InstalledProgramsList[2] = currentLine;
                                    InstalledProgramsList[2] = rgx.Replace(InstalledProgramsList[2], replacement);
                                }

                                if (currentLine.Contains("Apple Application Support (32-bit)"))
                                {
                                    currentLine = currentLine.Insert(34, ",");
                                    InstalledProgramsList[3] = currentLine;
                                    InstalledProgramsList[3] = rgx.Replace(InstalledProgramsList[3], replacement);
                                }

                                if (currentLine.Contains("Apple Mobile"))
                                {
                                    currentLine = currentLine.Insert(27, ",");
                                    InstalledProgramsList[4] = currentLine;
                                    InstalledProgramsList[4] = rgx.Replace(InstalledProgramsList[4], replacement);
                                }

                                if (currentLine.Contains("Bonjour"))
                                {
                                    currentLine = currentLine.Insert(7, ",");
                                    InstalledProgramsList[5] = currentLine;
                                    InstalledProgramsList[5] = rgx.Replace(InstalledProgramsList[5], replacement);
                                }

                                if (currentLine.Contains("Configuration Manager Client"))
                                {
                                    currentLine = currentLine.Insert(28, ",");
                                    InstalledProgramsList[6] = currentLine;
                                    InstalledProgramsList[6] = rgx.Replace(InstalledProgramsList[6], replacement);
                                }

                                if (currentLine.Contains("CrowdStrike"))
                                {
                                    currentLine = currentLine.Insert(26, ",");
                                    InstalledProgramsList[7] = currentLine;
                                    InstalledProgramsList[7] = rgx.Replace(InstalledProgramsList[7], replacement);
                                }

                                //CHROME + GOVERLAN MISSING FROM SPOTS 8/9

                                if (currentLine.Contains("iTunes"))
                                {
                                    currentLine = currentLine.Insert(6, ",");
                                    InstalledProgramsList[10] = currentLine;
                                    InstalledProgramsList[10] = rgx.Replace(InstalledProgramsList[10], replacement);
                                }

                                if (currentLine.Contains("Java 8 Update 171") && !currentLine.Contains("(64-bit)"))
                                {
                                    currentLine = currentLine.Insert(17, ",");
                                    InstalledProgramsList[11] = currentLine;
                                    InstalledProgramsList[11] = rgx.Replace(InstalledProgramsList[11], replacement);
                                }

                                if (currentLine.Contains("Java 8 Update 171 (64-bit)"))
                                {
                                    currentLine = currentLine.Insert(26, ",");
                                    InstalledProgramsList[12] = currentLine;
                                    InstalledProgramsList[12] = rgx.Replace(InstalledProgramsList[12], replacement);
                                }

                                if (currentLine.Contains("Microsoft Access MUI"))
                                {
                                    currentLine = currentLine.Insert(36, ",");
                                    InstalledProgramsList[13] = currentLine;
                                    InstalledProgramsList[13] = rgx.Replace(InstalledProgramsList[13], replacement);
                                }

                                if (currentLine.Contains("Microsoft Excel MUI"))
                                {
                                    currentLine = currentLine.Insert(34, ",");
                                    InstalledProgramsList[14] = currentLine;
                                    InstalledProgramsList[14] = rgx.Replace(InstalledProgramsList[14], replacement);
                                }

                                if (currentLine.Contains("Microsoft OneNote MUI"))
                                {
                                    currentLine = currentLine.Insert(36, ",");
                                    InstalledProgramsList[15] = currentLine;
                                    InstalledProgramsList[15] = rgx.Replace(InstalledProgramsList[15], replacement);
                                }

                                if (currentLine.Contains("Microsoft Outlook MUI"))
                                {
                                    currentLine = currentLine.Insert(36, ",");
                                    InstalledProgramsList[16] = currentLine;
                                    InstalledProgramsList[16] = rgx.Replace(InstalledProgramsList[16], replacement);
                                }

                                if (currentLine.Contains("Microsoft PowerPoint MUI"))
                                {
                                    currentLine = currentLine.Insert(39, ",");
                                    InstalledProgramsList[17] = currentLine;
                                    InstalledProgramsList[17] = rgx.Replace(InstalledProgramsList[17], replacement);
                                }

                                if (currentLine.Contains("Microsoft Publisher MUI"))
                                {
                                    currentLine = currentLine.Insert(38, ",");
                                    InstalledProgramsList[18] = currentLine;
                                    InstalledProgramsList[18] = rgx.Replace(InstalledProgramsList[18], replacement);
                                }

                                if (currentLine.Contains("Microsoft Skype for Business MUI"))
                                {
                                    currentLine = currentLine.Insert(47, ",");
                                    InstalledProgramsList[19] = currentLine;
                                    InstalledProgramsList[19] = rgx.Replace(InstalledProgramsList[19], replacement);
                                }

                                if (currentLine.Contains("Microsoft Word MUI"))
                                {
                                    currentLine = currentLine.Insert(31, ",");
                                    InstalledProgramsList[20] = currentLine;
                                    InstalledProgramsList[20] = rgx.Replace(InstalledProgramsList[20], replacement);
                                }

                                if (currentLine.Contains("Microsoft Silverlight"))
                                {
                                    currentLine = currentLine.Insert(21, ",");
                                    InstalledProgramsList[21] = currentLine;
                                    InstalledProgramsList[21] = rgx.Replace(InstalledProgramsList[21], replacement);
                                }

                                if (currentLine.Contains("Symantec Endpoint Protection"))
                                {
                                    currentLine = currentLine.Insert(28, ",");
                                    InstalledProgramsList[22] = currentLine;
                                    InstalledProgramsList[22] = rgx.Replace(InstalledProgramsList[22], replacement);
                                }

                                if (currentLine.Contains("NBC"))
                                {
                                    InstalledProgramsList[23] = currentLine;
                                    InstalledProgramsList[23] = rgx.Replace(InstalledProgramsList[23], replacement);
                                } // Doesnt Find NBC Offline Player

                            }

                            reader.Close();
                            File.Delete(@"c:\allprograms.txt");
                        }



                        //if (File.Exists(@"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe"))
                        //{
                        //ExecuteProgramsCommand(@"wmic datafile where name=""C:\\Program Files(x86)\\Google\\Chrome\\Application\\chrome.exe"" get Version /value >C:\chrome.txt");
                        //}
                        string FrameworkVersion = "";

                        try
                        {

                            RK = Registry.LocalMachine.OpenSubKey(@"SOFTWARE\Microsoft\NET Framework Setup\NDP\v3.5");
                            FrameworkVersion = RK.GetValue("Version").ToString();
                            RK.Close();

                            InstalledProgramsList[26] = "Microsoft .NET V3 Framework: " + "," + FrameworkVersion;
                        }
                        catch
                        {

                        }

                        try
                        {

                            RK = Registry.LocalMachine.OpenSubKey(@"SOFTWARE\Microsoft\NET Framework Setup\NDP\v4\Full");
                            FrameworkVersion = RK.GetValue("Version").ToString();
                            RK.Close();

                            InstalledProgramsList[25] = "Microsoft .NET V4 Framework: " + "," + FrameworkVersion;
                        }
                        catch (Exception)
                        {

                        }

                        //using (StreamWriter writer = File.CreateText(@"C:\CheckListApplications.csv"))
                        //{
                        //foreach (object obj in InstalledProgramsList)
                        //{
                        //  writer.WriteLine(obj);
                        //}

                        for (int i = 0; i <= 26; i++)
                        {
                            if (InstalledProgramsList[i] == "")
                            {
                                MissingProcessCount++;
                            }
                            stream.WriteLine(InstalledProgramsList[i]);
                        }
                        //}

                        if (MissingProcessCount > 0)
                        {
                            MessageBox.Show("Some programs could not be found. AppVerify will now open Apps & Features for further testing.");
                            ExecuteProgramsCommand("start ms-settings:appsfeatures");
                        }


                    }
                    catch
                    {
                        stream.WriteLine("Installed Applications could not be found. Please click the Installed Applications button.");
                    }
                }
            }

            FileInfo fi = new FileInfo(@"C:\AppVerify.csv");

            if (fi.Exists)
            {
                System.Diagnostics.Process.Start(@"C:\AppVerify.csv");
            }

            else
            {
                MessageBox.Show("Error: CSV File not Generated ");
            }

            //Application.Exit();

            KillCommandProcesses();

        }

        private void StudentVerify_Click(object sender, EventArgs e)
        {

            StringBuilder builder = new StringBuilder();
            int ProcessStartedCount = 0;
            //int NumberOfProcesses = 18;
            string path = @"C:\Student_App_Verification.txt";

            if (File.Exists(path))
            {
                File.Delete(path);
            }

            try
            {
                Process.Start(@"C:\Windows\SysWOW64\Macromed\Flash\FlashUtil_ActiveX.exe");
                builder.AppendLine("Adobe Shockwave Player : YES");
                ProcessStartedCount++;
            }
            catch
            {
                builder.AppendLine("Adobe Shockwave Player : NO, ADOBE PLAYER COULD NOT BE FOUND");
            }        

            try
            {
                Process.Start(@"C:\Program Files (x86)\Adobe\Acrobat Reader DC\Reader\AcroRd32.exe");
                builder.AppendLine("Adobe Reader DC MUI : YES");
                ProcessStartedCount++;
            }
            catch
            {
                //MessageBox.Show("Google CHROME could not be found", "Google CHROME Error");
                builder.AppendLine("Adobe Reader DC MUI : NO, ADOBE READER COULD NOT BE FOUND");
            }

            try
            {
                Process.Start(@"C:\Program Files\CrowdStrike\CSFalconService.exe");
                builder.AppendLine("CrowdStrike : YES");
                ProcessStartedCount++;
            }
            catch
            {
                builder.AppendLine("CrowdStrike : NO, CROWDSTRIKE COULD NOT BE FOUND");
            }

            try
            {
                Process.Start(@"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe");
                builder.AppendLine("Google Chrome : YES");
                ProcessStartedCount++;
            }
            catch
            {
                builder.AppendLine("Google Chrome : NO, CHROME COULD NOT BE FOUND");
            }

            try
            {
                Process.Start(@"C:\Program Files\iTunes\iTunes.exe");
                builder.AppendLine("Itunes : YES");
                ProcessStartedCount++;
            }
            catch
            {
                builder.AppendLine("Itunes : NO, ITUNES COULD NOT BE FOUND");
                // ITUNES
            }

            try
            {
                Process.Start(@"c:\Program Files (x86)\Java\jre1.8.0_171\bin\javaw.exe");
                builder.AppendLine("Java 8 32-bit : YES");
                ProcessStartedCount++;
            }
            catch
            {
                builder.AppendLine("Java 8 x32 : NO, JAVA 8 x32 COULD NOT BE FOUND");
                //JAVA 8 (32BIT)
            }

            try
            {
                Process.Start(@"c:\Program Files\Java\jre1.8.0_171\bin\javaw.exe");
                builder.AppendLine("Java 8 64-bit : YES");
                ProcessStartedCount++;
            }
            catch
            {
                builder.AppendLine("Java 8 x64 : NO, JAVA 8 x64 COULD NOT BE FOUND");
                //JAVA 8 (64BIT)
            }

            try
            {
                Process.Start(@"C:\Program Files (x86)\Microsoft Office\Office16\WINWORD.EXE");
                builder.AppendLine("Microsoft Word : YES");
                ProcessStartedCount++;
            }
            catch
            {
                builder.AppendLine("Microsoft Word : NO, WORD COULD NOT BE FOUND");
            }

            try
            {
                Process.Start(@"C:\Program Files (x86)\Microsoft Office\Office16\ONENOTE.EXE");
                builder.AppendLine("Microsoft OneNote : YES");
                ProcessStartedCount++;
            }
            catch
            {
                builder.AppendLine("Microsoft OneNote : NO, ONENOTE COULD NOT BE FOUND");
            }

            try
            {
                Process.Start(@"C:\Program Files (x86)\Microsoft Office\Office16\EXCEL.EXE");
                builder.AppendLine("Microsoft Excel : YES");
                ProcessStartedCount++;
            }
            catch
            {
                builder.AppendLine("Microsoft Excel : NO, EXCEL COULD NOT BE FOUND");
            }

            try
            {
                Process.Start(@"C:\Program Files (x86)\Microsoft Office\Office16\MSPUB.EXE");
                builder.AppendLine("Microsoft Publisher : YES");
                ProcessStartedCount++;
            }
            catch
            {
                builder.AppendLine("Microsoft Publisher : NO, PUBLISHER COULD NOT BE FOUND");
            }

            try
            {
                Process.Start(@"C:\Program Files (x86)\Microsoft Office\Office16\MSACCESS.EXE");
                builder.AppendLine("Microsoft Access : YES");
                ProcessStartedCount++;
            }
            catch
            {
                builder.AppendLine("Microsoft Access : NO, ACCESS COULD NOT BE FOUND");
            }

            try
            {
                Process.Start(@"C:\Program Files (x86)\Microsoft Office\Office16\OUTLOOK.EXE");
                builder.AppendLine("Microsoft Outlook : YES");
                ProcessStartedCount++;
            }
            catch
            {
                builder.AppendLine("Microsoft Outlook : NO, OUTLOOK COULD NOT BE FOUND");
            }

            try
            {
                Process.Start(@"C:\Program Files (x86)\Microsoft Office\Office16\POWERPNT.EXE");
                builder.AppendLine("Microsoft Powerpoint : YES");
                ProcessStartedCount++;
            }
            catch
            {
                builder.AppendLine("Microsoft Powerpoint : NO, POWERPOINT COULD NOT BE FOUND");
            }

            try
            {
                Process.Start(@"C:\Program Files (x86)\Microsoft Office\Office16\lync.exe");
                builder.AppendLine("Skype for Business : YES");
                ProcessStartedCount++;
            }
            catch
            {
                builder.AppendLine("Skype for Business : NO, SKYPE COULD NOT BE FOUND");
            }

            try
            {
                string temp = "";
                string SilverlightVersion = getVersionAutomatically(temp);
                //MessageBox.Show(SilverlightVersion);
                Process.Start(@"C:\Program Files (x86)\Microsoft Silverlight\"+ SilverlightVersion +@"\Silverlight.Configuration.exe");
                builder.AppendLine("Microsoft Silverlight : YES");
                ProcessStartedCount++;
            }
            catch
            {
                builder.AppendLine("Microsoft Silverlight : NO, MICROSOFT SILVERLIGHT COULD NOT BE FOUND");
                //MICROSOFT SILVERLIGHT
            }

            string SymantecPath = "";
            
            try
            {
                ExecuteCommand(@"wmic process where ""name = 'ccSvcHst.exe'"" get ExecutablePath >C:\symantec.txt");
                System.Threading.Thread.Sleep(5000);
                using (StreamReader reader = File.OpenText(@"C:\symantec.txt"))
                {
                    string currentLine;
                    while ((currentLine = reader.ReadLine()) != null)
                    {
                        if (currentLine.Contains(@"C:\"))
                        {
                            SymantecPath = currentLine;
                        }
                    }
                }

                File.Delete(@"C:\symantec.txt");

                SymantecPath = SymantecPath.Remove(83, 15);
                SymantecPath = SymantecPath +@"\SymCorpUI.exe";
                //MessageBox.Show(SymantecPath);

                //Process.Start(@"C:\Program Files (x86)\Symantec\Symantec Endpoint Protection\14.0.3897.1101.105\Bin\SymCorpUI.exe");
                Process.Start(SymantecPath);
                builder.AppendLine("Symantec : YES");
                ProcessStartedCount++;
            }
            catch
            {
                builder.AppendLine("Symantec : NO, SYMANTEC COULD NOT BE FOUND");
                //SYMANTEC ENDPOINT PROTECTION
            }

            try
            {
                Process.Start(@"C:\Program Files (x86)\NBC Universal, LLC\NBC Learn Offline\NBC Learn Offline.exe");
                builder.AppendLine("NBC Learn Offline: YES");
                ProcessStartedCount++;
            }
            catch
            {
                builder.AppendLine("NBC Learn Offline: NO, NBC LEARN OFFLINE COULD NOT BE FOUND");
                //NBC OFFLINE PLAYER
            }

            //builder.AppendLine("--------------------------------------------");
            //builder.AppendLine((ProcessStartedCount/NumberOfProcesses)*100 + "%" + " of processes started successfully!");

            if (!File.Exists(path))
            {
                using (StreamWriter SW = File.CreateText(path))
                {
                    SW.WriteLine("Application name : Did the application open?");
                    SW.WriteLine("--------------------------------------------");
                    SW.Write(builder.ToString());
                }
            }

            FileInfo fi = new FileInfo(path);

            if (fi.Exists)
            {
                System.Diagnostics.Process.Start(path);
            }

            System.Threading.Thread.Sleep(15000);

            DialogResult EndProcessesDialog = MessageBox.Show(new Form { TopMost = true },"Do you want to close all generated processes?", "End Student Verification", MessageBoxButtons.YesNo);

            if (EndProcessesDialog == DialogResult.Yes)
            {
                Process[] AllProcesses = Process.GetProcesses();

                foreach (var process in AllProcesses)
                {
                    if (process.MainWindowTitle != "")
                    {
                        string Process = process.ProcessName.ToUpper();
                        if (Process == "WINWORD")
                        {
                            process.Kill();
                        }
                        if (Process == "LYNC")
                        {
                            process.Kill();
                        }
                        if (Process == "OUTLOOK")
                        {
                            process.Kill();
                        }
                        if (Process == "MSPUB")
                        {
                            process.Kill();
                        }
                        if (Process == "MSACCESS")
                        {
                            process.Kill();
                        }
                        if (Process == "ONENOTE")
                        {
                            process.Kill();
                        }
                        if (Process == "EXCEL")
                        {
                            process.Kill();
                        }
                        if (Process == "CHROME")
                        {
                            process.Kill();
                        }
                        if (Process == "POWERPNT")
                        {
                            process.Kill();
                        }
                        if (Process == "ACRORD32")
                        {
                            process.Kill();
                        }
                        if (Process == "FLASHUTIL_ACTIVEX")
                        {
                            process.Kill();
                        }
                        if (Process == "NBC LEARN OFFLINE")
                        {
                            process.Kill();
                        }
                        if (Process == "ITUNES")
                        {
                            process.Kill();
                        }
                        if (Process == "JAVAW")
                        {
                            process.Kill();
                        }
                        if (Process == "SILVERLIGHT.CONFIGURATION")
                        {
                            process.Kill();
                        }
                        if (Process == "SYMCORPUI")
                        {
                            process.Kill();
                        }

                    }
                }
            }

            //Application.Exit();
            KillCommandProcesses();
        }

        private void TeacherVerify_Click(object sender, EventArgs e)
        {
            StringBuilder builder = new StringBuilder();
            int ProcessStartedCount = 0;
            //int NumberOfProcesses = 10;
            string path = @"C:\Teacher_App_Verification.txt";

            if (File.Exists(path)) {
                File.Delete(path);
            }

            try
            {
                Process.Start(@"C:\Program Files\CrowdStrike\CSFalconService.exe");
                builder.AppendLine("CrowdStrike : YES");
                ProcessStartedCount++;
            }
            catch
            {
                builder.AppendLine("CrowdStrike : NO, CROWDSTRIKE COULD NOT BE FOUND");
            }

            try
            {
                Process.Start(@"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe");
                builder.AppendLine("Google Chrome : YES");
                ProcessStartedCount++;
            }
            catch
            {
                builder.AppendLine("Google Chrome : NO, CHROME COULD NOT BE FOUND");
            }

            try
            {
                Process.Start(@"C:\Program Files (x86)\Microsoft Office\Office16\WINWORD.EXE");
                builder.AppendLine("Microsoft Word : YES");
                ProcessStartedCount++;
            }
            catch
            {
                builder.AppendLine("Microsoft Word : NO, WORD COULD NOT BE FOUND");
            }

            try
            {
                Process.Start(@"C:\Program Files (x86)\Microsoft Office\Office16\ONENOTE.EXE");
                builder.AppendLine("Microsoft OneNote : YES");
                ProcessStartedCount++;
            }
            catch
            {
                builder.AppendLine("Microsoft OneNote : NO, ONENOTE COULD NOT BE FOUND");
            }

            try
            {
                Process.Start(@"C:\Program Files (x86)\Microsoft Office\Office16\EXCEL.EXE");
                builder.AppendLine("Microsoft Excel : YES");
                ProcessStartedCount++;
            }
            catch
            {
                builder.AppendLine("Microsoft Excel : NO, EXCEL COULD NOT BE FOUND");
            }

            try
            {
                Process.Start(@"C:\Program Files (x86)\Microsoft Office\Office16\MSPUB.EXE");
                builder.AppendLine("Microsoft Publisher : YES");
                ProcessStartedCount++;
            }
            catch
            {
                builder.AppendLine("Microsoft Publisher : NO, PUBLISHER COULD NOT BE FOUND");
            }

            try
            {
                Process.Start(@"C:\Program Files (x86)\Microsoft Office\Office16\MSACCESS.EXE");
                builder.AppendLine("Microsoft Access : YES");
                ProcessStartedCount++;
            }
            catch
            {
                builder.AppendLine("Microsoft Access : NO, ACCESS COULD NOT BE FOUND");
            }

            try
            {
                Process.Start(@"C:\Program Files (x86)\Microsoft Office\Office16\OUTLOOK.EXE");
                builder.AppendLine("Microsoft Outlook : YES");
                ProcessStartedCount++;
            }
            catch
            {
                builder.AppendLine("Microsoft Outlook : NO, OUTLOOK COULD NOT BE FOUND");
            }

            try
            {
                Process.Start(@"C:\Program Files (x86)\Microsoft Office\Office16\POWERPNT.EXE");
                builder.AppendLine("Microsoft Powerpoint : YES");
                ProcessStartedCount++;
            }
            catch
            {
                builder.AppendLine("Microsoft Powerpoint : NO, POWERPOINT COULD NOT BE FOUND");
            }

            try
            {
                Process.Start(@"C:\Program Files (x86)\Microsoft Office\Office16\lync.exe");
                builder.AppendLine("Skype for Business : YES");
                ProcessStartedCount++;
            }
            catch
            {
                builder.AppendLine("Skype for Business : NO, SKYPE COULD NOT BE FOUND");
            }

            //int Percentage = (ProcessStartedCount / NumberOfProcesses) * 100;

            //builder.AppendLine("--------------------------------------------");
            //builder.AppendLine(Percentage + "%" + " of processes started successfully!");

            if (!File.Exists(path))
            {
                using (StreamWriter SW = File.CreateText(path))
                {
                    SW.WriteLine("Application name : Did the application open?");
                    SW.WriteLine("--------------------------------------------");
                    SW.Write(builder.ToString());
                }
            }

            FileInfo fi = new FileInfo(path);

            if (fi.Exists)
            {
                System.Diagnostics.Process.Start(path);
            }

            System.Threading.Thread.Sleep(5000);

            DialogResult EndProcessesDialog = MessageBox.Show(new Form { TopMost = true }, "Do you want to close all generated processes?", "End Teacher Verification", MessageBoxButtons.YesNo);

            if (EndProcessesDialog == DialogResult.Yes)
            {
                Process[] AllProcesses = Process.GetProcesses();

                foreach (var process in AllProcesses)
                {
                    if (process.MainWindowTitle != "")
                    {
                        string Process = process.ProcessName.ToUpper();
                        if (Process == "WINWORD")
                        {
                            process.Kill();
                        }
                        if (Process == "LYNC")
                        {
                            process.Kill();
                        }
                        if (Process == "OUTLOOK")
                        {
                            process.Kill();
                        }
                        if (Process == "MSPUB")
                        {
                            process.Kill();
                        }
                        if (Process == "MSACCESS")
                        {
                            process.Kill();
                        }
                        if (Process == "ONENOTE")
                        {
                            process.Kill();
                        }
                        if (Process == "EXCEL")
                        {
                            process.Kill();
                        }
                        if (Process == "CHROME")
                        {
                            process.Kill();
                        }
                        if (Process == "POWERPNT")
                        {
                            process.Kill();
                        }
                        
                    }
                }
            }

            else if (EndProcessesDialog == DialogResult.No)
            {
                return;
            }

            KillCommandProcesses();

        }

        private void RemovedFiles_Click(object sender, EventArgs e)
        {
            string CSVpath = @"C:\AppVerify.csv";
            string Studentpath = @"C:\Student_App_Verification.txt";
            string Teacherpath = @"C:\Teacher_App_Verification.txt";
            string InstalledProgramsPath = @"C:\CheckListApplications.csv";

            int Pathcount = 0;

            Process[] AllProcesses = Process.GetProcesses();

            foreach (var process in AllProcesses)
            {
                if (process.MainWindowTitle != "")
                {
                    string Process = process.ProcessName.ToUpper();
                    if (Process == "NOTEPAD")
                    {
                        process.Kill();
                    }
                    if (Process == "EXCEL")
                    {
                        process.Kill();
                    }
                }
            }

            System.Threading.Thread.Sleep(3000);

            try
            {
                if (File.Exists(CSVpath))
                {
                    File.Delete(CSVpath);
                    Pathcount++;
                }
            }
            catch
            {
                MessageBox.Show("Please close the CSV file and try again!");
            }

            if (File.Exists(Studentpath))
            {
                File.Delete(Studentpath);
                Pathcount++;
            }

            if (File.Exists(Teacherpath))
            {
                File.Delete(Teacherpath);
                Pathcount++;
            }

            if (File.Exists(InstalledProgramsPath))
            {
                File.Delete(InstalledProgramsPath);
            }

            if (Pathcount == 0)
            {
                MessageBox.Show("AppVerify generated files have already been removed!");
            }
            else
            {
                MessageBox.Show("AppVerify generated files have been removed!");
            }

            KillCommandProcesses();
        }

        private void InstalledApps_Click(object sender, EventArgs e)
        {
            int MissingProcessCount = 0;
            string pattern = "\\s+";
            string replacement = " ";
            Regex rgx = new Regex(pattern);
            //version = rgx.Replace(version, replacement);

            Process[] AllProcesses = Process.GetProcesses();

            foreach (var process in AllProcesses)
            {
                if (process.MainWindowTitle != "")
                {
                    string Process = process.ProcessName.ToUpper();
                    if (Process == "Excel".ToUpper())
                    {
                        process.Kill();
                    }
                }
            }

            if (File.Exists(@"C:\chrome.txt"))
            {
                File.Delete(@"C:\chrome.txt");
            }

            if (File.Exists(@"C:\allprograms.txt"))
            {
                File.Delete(@"C:\allprograms.txt");
            }

            if (File.Exists(@"C:\CheckListApplications.csv"))
            {
                File.Delete(@"C:\CheckListApplications.csv");
            }

            string CommandInstalledApps = @"wmic product get name, version >C:\allprograms.txt";
            ExecuteProgramsCommand(CommandInstalledApps);

            String[] InstalledProgramsList = new String[27];

            for (int i = 0; i <= 26; i++)
            {
                InstalledProgramsList[i] = "";
            }

            InstalledProgramsList[24] = "------- Program Features Not Listed -------";

            System.Threading.Thread.Sleep(40000);

            using (StreamReader reader = File.OpenText(@"c:\allprograms.txt"))
            {
                string currentLine;
                while ((currentLine = reader.ReadLine()) != null)
                {

                    if (currentLine.Contains("Adobe AIR"))
                    {
                        currentLine = currentLine.Insert(9, ",");
                        InstalledProgramsList[0] = currentLine;
                        InstalledProgramsList[0] = rgx.Replace(InstalledProgramsList[0], replacement);
                    }

                    if (currentLine.Contains("Adobe Acrobat"))
                    {
                        currentLine = currentLine.Insert(23, ",");
                        InstalledProgramsList[1] = currentLine;
                        InstalledProgramsList[1] = rgx.Replace(InstalledProgramsList[1], replacement);
                    }

                    if (currentLine.Contains("Adobe Shockwave"))
                    {
                        //currentLine = currentLine.Insert(23, ",");
                        InstalledProgramsList[2] = currentLine;
                        InstalledProgramsList[2] = rgx.Replace(InstalledProgramsList[2], replacement);
                    }

                    if (currentLine.Contains("Apple Application Support (32-bit)"))
                    {
                        currentLine = currentLine.Insert(34, ",");
                        InstalledProgramsList[3] = currentLine;
                        InstalledProgramsList[3] = rgx.Replace(InstalledProgramsList[3], replacement);
                    }

                    if (currentLine.Contains("Apple Mobile"))
                    {
                        currentLine = currentLine.Insert(27, ",");
                        InstalledProgramsList[4] = currentLine;
                        InstalledProgramsList[4] = rgx.Replace(InstalledProgramsList[4], replacement);
                    }

                    if (currentLine.Contains("Bonjour"))
                    {
                        currentLine = currentLine.Insert(7, ",");
                        InstalledProgramsList[5] = currentLine;
                        InstalledProgramsList[5] = rgx.Replace(InstalledProgramsList[5], replacement);
                    }

                    if (currentLine.Contains("Configuration Manager Client"))
                    {
                        currentLine = currentLine.Insert(28, ",");
                        InstalledProgramsList[6] = currentLine;
                        InstalledProgramsList[6] = rgx.Replace(InstalledProgramsList[6], replacement);
                    }
                    
                    if (currentLine.Contains("CrowdStrike"))
                    {
                        currentLine = currentLine.Insert(26, ",");
                        InstalledProgramsList[7] = currentLine;
                        InstalledProgramsList[7] = rgx.Replace(InstalledProgramsList[7], replacement);
                    }

                    //CHROME + GOVERLAN MISSING FROM SPOTS 8/9

                    if (currentLine.Contains("iTunes"))
                    {
                        currentLine = currentLine.Insert(6, ",");
                        InstalledProgramsList[10] = currentLine;
                        InstalledProgramsList[10] = rgx.Replace(InstalledProgramsList[10], replacement);
                    }

                    if (currentLine.Contains("Java 8 Update 171") && !currentLine.Contains("(64-bit)"))
                    {
                        currentLine = currentLine.Insert(17, ",");
                        InstalledProgramsList[11] = currentLine;
                        InstalledProgramsList[11] = rgx.Replace(InstalledProgramsList[11], replacement);
                    }

                    if (currentLine.Contains("Java 8 Update 171 (64-bit)"))
                    {
                        currentLine = currentLine.Insert(26, ",");
                        InstalledProgramsList[12] = currentLine;
                        InstalledProgramsList[12] = rgx.Replace(InstalledProgramsList[12], replacement);
                    }

                    if (currentLine.Contains("Microsoft Access MUI"))
                    {
                        currentLine = currentLine.Insert(36, ",");
                        InstalledProgramsList[13] = currentLine;
                        InstalledProgramsList[13] = rgx.Replace(InstalledProgramsList[13], replacement);
                    }

                    if (currentLine.Contains("Microsoft Excel MUI"))
                    {
                        currentLine = currentLine.Insert(34, ",");
                        InstalledProgramsList[14] = currentLine;
                        InstalledProgramsList[14] = rgx.Replace(InstalledProgramsList[14], replacement);
                    }

                    if (currentLine.Contains("Microsoft OneNote MUI"))
                    {
                        currentLine = currentLine.Insert(36, ",");
                        InstalledProgramsList[15] = currentLine;
                        InstalledProgramsList[15] = rgx.Replace(InstalledProgramsList[15], replacement);
                    }

                    if (currentLine.Contains("Microsoft Outlook MUI"))
                    {
                        currentLine = currentLine.Insert(36, ",");
                        InstalledProgramsList[16] = currentLine;
                        InstalledProgramsList[16] = rgx.Replace(InstalledProgramsList[16], replacement);
                    }

                    if (currentLine.Contains("Microsoft PowerPoint MUI"))
                    {
                        currentLine = currentLine.Insert(39, ",");
                        InstalledProgramsList[17] = currentLine;
                        InstalledProgramsList[17] = rgx.Replace(InstalledProgramsList[17], replacement);
                    }

                    if (currentLine.Contains("Microsoft Publisher MUI"))
                    {
                        currentLine = currentLine.Insert(38, ",");
                        InstalledProgramsList[18] = currentLine;
                        InstalledProgramsList[18] = rgx.Replace(InstalledProgramsList[18], replacement);
                    }

                    if (currentLine.Contains("Microsoft Skype for Business MUI"))
                    {
                        currentLine = currentLine.Insert(47, ",");
                        InstalledProgramsList[19] = currentLine;
                        InstalledProgramsList[19] = rgx.Replace(InstalledProgramsList[19], replacement);
                    }

                    if (currentLine.Contains("Microsoft Word MUI"))
                    {
                        currentLine = currentLine.Insert(31, ",");
                        InstalledProgramsList[20] = currentLine;
                        InstalledProgramsList[20] = rgx.Replace(InstalledProgramsList[20], replacement);
                    }

                    if (currentLine.Contains("Microsoft Silverlight"))
                    {
                        currentLine = currentLine.Insert(21, ",");
                        InstalledProgramsList[21] = currentLine;
                        InstalledProgramsList[21] = rgx.Replace(InstalledProgramsList[21], replacement);
                    }

                    if (currentLine.Contains("Symantec Endpoint Protection"))
                    {
                        currentLine = currentLine.Insert(28, ",");
                        InstalledProgramsList[22] = currentLine;
                        InstalledProgramsList[22] = rgx.Replace(InstalledProgramsList[22], replacement);
                    }

                    if (currentLine.Contains("NBC"))
                    {
                        InstalledProgramsList[23] = currentLine;
                        InstalledProgramsList[23] = rgx.Replace(InstalledProgramsList[23], replacement);
                    } // Doesnt Find NBC Offline Player
                    
                }

                reader.Close();
                File.Delete(@"c:\allprograms.txt");
            }

            //if (File.Exists(@"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe"))
            //{
            //ExecuteProgramsCommand(@"wmic datafile where name=""C:\\Program Files(x86)\\Google\\Chrome\\Application\\chrome.exe"" get Version /value >C:\chrome.txt");
            //}
            string FrameworkVersion = "";

            try
            {
                RegistryKey RK;
                RK = Registry.LocalMachine.OpenSubKey(@"SOFTWARE\Microsoft\NET Framework Setup\NDP\v3.5");
                FrameworkVersion = RK.GetValue("Version").ToString();
                RK.Close();

                InstalledProgramsList[26] = "Microsoft .NET V3 Framework: " + "," + FrameworkVersion;
            }
            catch
            {

            }

            try
            {
                RegistryKey RK;
                RK = Registry.LocalMachine.OpenSubKey(@"SOFTWARE\Microsoft\NET Framework Setup\NDP\v4\Full");
                FrameworkVersion = RK.GetValue("Version").ToString();
                RK.Close();

                InstalledProgramsList[25] = "Microsoft .NET V4 Framework: " + "," + FrameworkVersion;
            }
            catch (Exception)
            {
                
            }
           
            using (StreamWriter writer = File.CreateText(@"C:\CheckListApplications.csv"))
            {
                //foreach (object obj in InstalledProgramsList)
                //{
                  //  writer.WriteLine(obj);
                //}

                for (int i = 0; i <= 26; i++)
                {
                    if (InstalledProgramsList[i] == "")
                    {
                        MissingProcessCount++;
                    }
                    writer.WriteLine(InstalledProgramsList[i]);
                }
            }

            if (MissingProcessCount > 0)
            {
                MessageBox.Show("Some programs could not be found. AppVerify will now open Apps & Features for further testing.");
                ExecuteProgramsCommand("start ms-settings:appsfeatures");
            }

            System.Threading.Thread.Sleep(2000);

            FileInfo fi = new FileInfo(@"C:\CheckListApplications.csv");

            if (fi.Exists)
            {
                System.Diagnostics.Process.Start(@"C:\CheckListApplications.csv");
            }

            KillCommandProcesses();
        }
        
        private void About_Click(object sender, EventArgs e)
        {
            string Version = Assembly.GetExecutingAssembly().GetName().Version.ToString();
            string Year = DateTime.Now.Year.ToString();
            string Date = DateTime.Now.ToString("MM/dd/yyyy");
            MessageBox.Show("Created by: James Goodman, Rahul Batra, Moaad Hidias, Antonio Smallwood and Keith Sue." + "\r\n" + "----------------------------------------------------------------------------------" + "\r\n" + "Application Version as of " + Date + ": " + Version + "\r\n" + "----------------------------------------------------------------------------------" + "\r\n" + "CopyRight © SD&E " + Year, "AppVerify");
        }
    }

}
