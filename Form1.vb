Imports VB = Microsoft.VisualBasic
Public Class Form1
    Public Sub wait(ByVal seconds As Single)
        Static start As Single
        start = VB.Timer()
        Do While VB.Timer() < start + seconds
            System.Windows.Forms.Application.DoEvents()
        Loop
    End Sub

    Private Sub Button1_Click_1(sender As Object, e As EventArgs) Handles Button1.Click
        Dim result As Integer = MessageBox.Show("Please close all Internet Browsers for successful automation then press OK.",
        "Close Internet Browsers", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
        If result = DialogResult.Yes Then
            Process.Start("rundll32.exe", "shell32.dll,Control_RunDLL inetcpl.cpl")
            wait(1)
            SendKeys.Send("google.com")
            SendKeys.Send("{TAB}")
            SendKeys.Send("{TAB}")
            SendKeys.Send("{TAB}")
            SendKeys.Send("{TAB}")
            SendKeys.Send("{TAB}")
            SendKeys.Send("{TAB}")
            SendKeys.Send("{TAB}")
            SendKeys.Send("{TAB}")
            SendKeys.Send("{TAB}")
            SendKeys.Send("{TAB}")
            SendKeys.Send("{TAB}")
            SendKeys.Send("{TAB}")
            SendKeys.Send("{TAB}")
            SendKeys.Send("{TAB}")
            SendKeys.Send("{TAB}")
            SendKeys.Send("{RIGHT}")
            SendKeys.Send("{RIGHT}")
            SendKeys.Send("{RIGHT}")
            SendKeys.Send("{RIGHT}")
            SendKeys.Send("{RIGHT}")
            SendKeys.Send("{RIGHT}")
            SendKeys.Send("{TAB}")
            SendKeys.Send("{TAB}")
            SendKeys.Send("{ENTER}")
            SendKeys.Send("{TAB}")
            SendKeys.Send("{ENTER}")
            SendKeys.Send("{LEFT}")
            SendKeys.Send("{ENTER}")
        End If

    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        wait(1)
        Dim process As New Process()
        process.StartInfo.FileName = "cmd"
        process.StartInfo.Verb = "runas"
        process.StartInfo.UseShellExecute = False
        process.Start()
        wait(1)
        SendKeys.Send("{t}")
        SendKeys.Send("{a}")
        SendKeys.Send("{s}")
        SendKeys.Send("{k}")
        SendKeys.Send("{m}")
        SendKeys.Send("{g}")
        SendKeys.Send("{r}")
        SendKeys.Send("{.}")
        SendKeys.Send("{e}")
        SendKeys.Send("{x}")
        SendKeys.Send("{e}")
        SendKeys.Send("{ENTER}")
        wait(1)
        process.Kill()
        MessageBox.Show("
***Visit the Startup tab on the Task Manager***

Disable non-critical programs from your computers startup.

Avoid disabling startup programs: Nvidia, Realtek, Audio Manager, Wifi Manager, Intel, Windows.

If you lose Audio or Internet after restarting, you will need to re-enable it in the startup menu and restart.

Disabled startup programs are still 100% functional when opened manually with a double click, so no need to worry.

Disabling non-critical startup programs will improve computer performance.

This process is not fully automated due to Windows internal security features.",
        "Startup Programs Guide", MessageBoxButtons.OK, MessageBoxIcon.Information)
    End Sub

    Private Sub Button4_Click_1(sender As Object, e As EventArgs) Handles Button4.Click
        Dim process As New Process()
        process.StartInfo.FileName = "cmd.exe "
        process.StartInfo.Verb = "runas"
        process.StartInfo.UseShellExecute = True
        process.Start()
    End Sub

    Private Sub Button7_Click_1(sender As Object, e As EventArgs) Handles Button7.Click
        MessageBox.Show("Read carefully

Birds Eye View Optimization Tool is created by leading Windows 10 developers and is intended to bring the power of a computer technician by combining automation and guides. 

Computer speed will improve dramatically after completing each step and restarting the computer once or twice through the process.

If you choose to use an automation tool in this program please do not use your keyboard and mouse until the chosen tool is finished.

In order to fully optimize your computer follow the directions presented in the message boxes to complete each step.

In the event that User Account Control asks for your permission, press Yes to continue.

Birds Eye View Optimization Tool is designed for Windows 10.
",
        "Read Me", MessageBoxButtons.OK, MessageBoxIcon.Information)
    End Sub


    Private Sub Button8_Click(sender As Object, e As EventArgs)
        MessageBox.Show(".

If you recieve an error that does not allow you to complete the 'sfc /scannow' function, use the Command Prompt (Admin) tool and type in 
'DISM /Online /Cleanup-Image /RestoreHealth'
When it is complete try the 'sfc /scannow' function.

If both of these fail to restore your operating system, you may need to use the Command Prompt (Admin) tool and type in 'chkdsk /r', restart the computer, then try the 'sfc /scannow' function.

If all of the above fail, and the tools provided on this application fail to fix the 'sfc /scannow' function, you may need to reset your copy of Windows 10.",
        "Read Me", MessageBoxButtons.OK, MessageBoxIcon.Information)
    End Sub


    Private Sub Button9_Click(sender As Object, e As EventArgs) Handles Button9.Click
        MessageBox.Show("Uninstall programs that cause disturbances and pop-ups on your computer.

Be sure to have a maximum of one antivirus on your computer as any extra antiviruses will cause system instability.

Caution: Do not uninstall system drivers. Example: Microsoft, Windows, Nvidia, HD Audio, AMD.

Press Ok to uninstall programs with instructions.",
        "Uninstalling Programs", MessageBoxButtons.OK, MessageBoxIcon.Information)
        Process.Start("appwiz.cpl")
        MessageBox.Show("Uninstall programs that cause disturbances and pop-ups on your computer.

Be sure to have a maximum of one antivirus on your computer as any extra antiviruses will cause system instability.

Caution: Do not uninstall system drivers. Example: Microsoft, Windows, Nvidia, HD Audio, AMD, Java",
        "Uninstalling Programs", MessageBoxButtons.OK, MessageBoxIcon.Information)
    End Sub


    Private Sub SuggestionsComingSoonToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles SuggestionsComingSoonToolStripMenuItem.Click
        MessageBox.Show("We are happy to hear from you. 

Contact us at admin@birdseyeviewtech.com",
        "Contact Us", MessageBoxButtons.OK, MessageBoxIcon.Information)
    End Sub

    Private Sub Button16_Click(sender As Object, e As EventArgs)
        Process.Start("rundll32.exe", "shell32.dll,Control_RunDLL")
        wait(3)
        SendKeys.Send("{TAB}")
        SendKeys.Send("{TAB}")
        SendKeys.Send("{TAB}")
        SendKeys.Send("{TAB}")
        SendKeys.Send("^{F}")

    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        MessageBox.Show("Make sure Caps Lock is off then press OK to begin automation",
       "Caps Lock: Off", MessageBoxButtons.OK, MessageBoxIcon.Information)

        Process.Start("chrome.exe")
        Dim p() As Process
        p = Process.GetProcessesByName("chrome.exe")
        If p.Count > 0 Then
            wait(3)
            SendKeys.Send("chrome://settings/resetProfileSettings?origin=userclick&search=reset")
        Else
            wait(3)
            SendKeys.Send("chrome://settings/resetProfileSettings?origin=userclick&search=reset")
            SendKeys.Send("{ENTER}")
            wait(3)
            SendKeys.Send("{TAB}")
            SendKeys.Send("{TAB}")
            SendKeys.Send("{TAB}")
            SendKeys.Send("{ENTER}")
        End If
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        MessageBox.Show("Enable useful extensions that were disabled during the automation process.",
       "Enable Useful Extensions", MessageBoxButtons.OK, MessageBoxIcon.Information)
        Process.Start("chrome.exe")
        Dim r() As Process
        r = Process.GetProcessesByName("chrome.exe")
        If r.Count > 0 Then
            wait(1)
            SendKeys.Send("chrome://extensions/")
        Else
            wait(1)
            SendKeys.Send("chrome://extensions/""{ENTER}")
        End If
    End Sub

    Private Sub Button10_Click(sender As Object, e As EventArgs) Handles Button10.Click
        Dim result As Integer = MessageBox.Show("Would you like to be directed to bleepingcomputers.com to download JRT (Junkware Removal Tool)?

This tool is excellent at identifying potentially unwanted programs.", "Download JRT", MessageBoxButtons.YesNo, MessageBoxIcon.Question)

        If result = DialogResult.Yes Then
            Process.Start("https://www.bleepingcomputer.com/download/junkware-removal-tool/")
        End If
    End Sub

    Private Sub Button11_Click(sender As Object, e As EventArgs) Handles Button11.Click
        Dim result As Integer = MessageBox.Show("Would you like to be directed to malwarebytes.com to download AdwCleaner?

This tool is excellent at identifying and removing adware from your system

Removing adware from your system could significantly improve performance", "Download AdwCleaner", MessageBoxButtons.YesNo, MessageBoxIcon.Question)

        If result = DialogResult.Yes Then
            Process.Start("https://www.malwarebytes.com/adwcleaner/")
        End If
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        Dim result As Integer = MessageBox.Show("Would you like to be directed to www.hitmanpro.com to download HitmanPro?

This tool is excellent at identifying and removing malware from your system.

Removing malware from your system could significantly improve performance.

The trail version is free for 30 days and doesn't require a credit card.

This program only works when you start a scan and does not include 24/7 protection.

", "Download HitmanPro", MessageBoxButtons.YesNo, MessageBoxIcon.Question)

        If result = DialogResult.Yes Then
            If Environment.Is64BitOperatingSystem Then
                Process.Start("https://www.hitmanpro.com/en-us/downloads.aspx")
                Process.Start("https://dl.surfright.nl/HitmanPro_x64.exe")
            Else
                Process.Start("https://www.hitmanpro.com/en-us/downloads.aspx")
                Process.Start("https://dl.surfright.nl/HitmanPro.exe")
            End If
        End If
    End Sub

    Private Sub Button13_Click(sender As Object, e As EventArgs) Handles Button13.Click
        Dim result As Integer = MessageBox.Show("Would you like to be directed to https://www.wisecleaner.com to download Wise Registry Cleaner?

This tool is excellent at identifying and removing bad registry keys from your system

Restoring the registry could significantly restore computer stability

Do you want to proceed?", "Download Wise Registry Cleaner", MessageBoxButtons.YesNo, MessageBoxIcon.Question)

        If result = DialogResult.Yes Then
            Process.Start("https://www.wisecleaner.com/wise-registry-cleaner.html")
            wait(3)
            MessageBox.Show("Step A: Install Wise Registry Cleaner
Step B: Initialize a Deep Scan
Step C: When the Deep Scan is done press Clean
Step D: Under the System Tune Up tab press Optimze
Step E: Uninstall Wise Registry Cleaner ",
        "Wise Registry Cleaner Guide", MessageBoxButtons.OK, MessageBoxIcon.Information)
        End If
    End Sub

    Private Sub Button14_Click(sender As Object, e As EventArgs) Handles Button14.Click
        Dim result As Integer = MessageBox.Show("If your welcome screen startup used to be a lot faster it is likely that your hard drive is failing.

Computer Hard Drives typically slow down drastically after about four years of daily use. 

If you believe this is the case then you will need to verify the integrety of your hard drive by downloading HDDScan.

See the guide available above to verify the integrity of your hard drive or visit HDDScan.com for instructions.

Would you like to be directed to the HDDScan download at http://hddscan.com/?
", "Download HDDScan", MessageBoxButtons.YesNo, MessageBoxIcon.Question)

        If result = DialogResult.Yes Then
            Process.Start("http://hddscan.com/")
        End If
    End Sub

    Private Sub Button12_Click(sender As Object, e As EventArgs) Handles Button12.Click
        wait(1)
        Dim process As New Process()
        process.StartInfo.FileName = "cmd"
        process.StartInfo.Verb = "runas"
        process.StartInfo.UseShellExecute = False
        process.Start()
        wait(2)
        SendKeys.Send("{c}")
        SendKeys.Send("{d}")
        SendKeys.Send(" ")
        SendKeys.Send("{%}")
        SendKeys.Send("{t}")
        SendKeys.Send("{e}")
        SendKeys.Send("{m}")
        SendKeys.Send("{p}")
        SendKeys.Send("{%}")
        SendKeys.Send("{ENTER}")
        wait(2)
        SendKeys.Send("{c}")
        SendKeys.Send("{d}")
        SendKeys.Send(" ")
        SendKeys.Send("{.}")
        SendKeys.Send("{.}")
        SendKeys.Send("{ENTER}")
        wait(2)
        SendKeys.Send("{r}")
        SendKeys.Send("{d}")
        SendKeys.Send(" ")
        SendKeys.Send("{T}")
        SendKeys.Send("{e}")
        SendKeys.Send("{m}")
        SendKeys.Send("{p}")
        SendKeys.Send(" ")
        SendKeys.Send("{/}")
        SendKeys.Send("{s}")
        SendKeys.Send(" ")
        SendKeys.Send("{ENTER}")
        wait(2)
        SendKeys.Send("{Y}")
        SendKeys.Send("{ENTER}")
        wait(2)
        SendKeys.Send("{m}")
        SendKeys.Send("{d}")
        SendKeys.Send(" ")
        SendKeys.Send("{T}")
        SendKeys.Send("{e}")
        SendKeys.Send("{m}")
        SendKeys.Send("{p}")
        SendKeys.Send("{ENTER}")
        wait(2)
        SendKeys.Send("(%{F4})")

    End Sub

    Private Sub Button15_Click(sender As Object, e As EventArgs)
        Process.Start("https://www.paypal.com/cgi-bin/webscr?cmd=_s-xclick&hosted_button_id=TWSZAM84TDBFC")
    End Sub

    Private Sub Button17_Click(sender As Object, e As EventArgs) Handles Button17.Click
        Process.Start("https://www.paypal.com/cgi-bin/webscr?cmd=_s-xclick&hosted_button_id=TWSZAM84TDBFC")
    End Sub

    Private Sub Button18_Click(sender As Object, e As EventArgs) Handles Button18.Click
        wait(1)
        Dim process As New Process()
        process.StartInfo.FileName = "cmd"
        process.StartInfo.Verb = "runas"
        process.StartInfo.UseShellExecute = False
        process.Start()
        wait(1)
        SendKeys.Send("{t}")
        SendKeys.Send("{a}")
        SendKeys.Send("{s}")
        SendKeys.Send("{k}")
        SendKeys.Send("{m}")
        SendKeys.Send("{g}")
        SendKeys.Send("{r}")
        SendKeys.Send("{.}")
        SendKeys.Send("{e}")
        SendKeys.Send("{x}")
        SendKeys.Send("{e}")
        SendKeys.Send("{ENTER}")
        wait(1)
        process.Kill()
    End Sub

    Private Sub Button16_Click_1(sender As Object, e As EventArgs) Handles Button16.Click
        Dim result As Integer = MessageBox.Show(" 'sfc /scannow' command corruption can lead to other issues. 
                         
You will probably fix the 'sfc /scannow' command with the 'DISM /Online /Cleanup-Image /RestoreHealth' command.
                         
System corruption is common in computers with old hard drives or computers that have been heavily infected with viruses.

Do you want to proceed?", "Instructions", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
        If result = DialogResult.Yes Then
            Dim process As New Process()
            process.StartInfo.FileName = "cmd.exe "
            process.StartInfo.Verb = "runas"
            process.StartInfo.UseShellExecute = True
            process.Start()
            wait(1)
            Clipboard.SetText("DISM /Online /Cleanup-Image /RestoreHealth")
            Clipboard.SetText("DISM /Online /Cleanup-Image /RestoreHealth")
            MessageBox.Show("Right click once or twice of the black box then press enter in order to start the 
'DISM /Online /Cleanup-Image /RestoreHealth' command.


", "Instructions", MessageBoxButtons.OK, MessageBoxIcon.Information)
        End If
    End Sub

    Private Sub Button19_Click(sender As Object, e As EventArgs) Handles Button19.Click
        Dim result As Integer = MessageBox.Show("Windows corruption is common, the 'sfc /scannow' command will attempt to scan and fix your system files
                         
Windows corruption is caused by failing hard drives or computers that have been heavily infected with viruses

Do you want to proceed?", "Instructions", MessageBoxButtons.YesNo, MessageBoxIcon.Question)

        If result = DialogResult.Yes Then
            Dim process As New Process()
            process.StartInfo.FileName = "cmd.exe "
            process.StartInfo.Verb = "runas"
            process.StartInfo.UseShellExecute = True
            process.Start()
            wait(1)
            Clipboard.SetText("sfc /scannow")
            Clipboard.SetText("sfc /scannow")
            MessageBox.Show("Right click once or twice inside of the black box to paste 
'sfc /scannow' then press enter in order to start the process.

If you have recieved an error use the (Repair SFC) tool",
                        "Instructions", MessageBoxButtons.OK, MessageBoxIcon.Information)
        End If
    End Sub

    Private Sub VersionToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles VersionToolStripMenuItem.Click
        MessageBox.Show("BEV Optimization Tool 1.1",
                        "Version", MessageBoxButtons.OK, MessageBoxIcon.Information)
    End Sub

    Private Sub PrinterGuidesComingSoonToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles PrinterGuidesComingSoonToolStripMenuItem.Click

    End Sub

    Private Sub DisclaimerToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles DisclaimerToolStripMenuItem.Click
        MessageBox.Show("Tools may be added or removed in each new update as we evaluate their relavence.

Birds Eye View Optimization is not paid to advertise programs.

We are not responsible for any data loss so be sure to back up your data before optimizing as failing hard drives could fail at any time.
",
                        "Information", MessageBoxButtons.OK, MessageBoxIcon.Information)
    End Sub
End Class