# Python 2.7 program 
# Remote Workstation health
# 11/29/2017
# by Matthew R. Plahuta

from subprocess import check_output
import wmi
import socket
import win32com.client
import os, time

class PcHealth:
    """
    This Class was designed to gather important workstation health information and will aid you in troubleshooting and 
    will result in a faster soulution for our clients.
    """
    def __init__(self, hostname):
        self.ava_code = { 1 : 'Other', 2 : 'Unknown', 3 : 'Running / Full Power', 4 : 'Warning', 5 : 'In Test', 6 : 'Not Applicable',
        7 : 'Power Off', 8 : 'Off Line', 9 : 'Off Duty', 10 : 'Degraded', 11 : 'Degraded', 12 : 'Install Error', 
        13 : 'Power Save - Unknown ', 14 : 'Power Save - Low Power Mode', 15 : 'Power Save - Standby', 16 : 'Power Cycle ', 17 : 
        'Power Save - Warning', 18 : 'Paused', 19 : 'Not Ready', 20 : 'Not Configured', 21 : 'Quiesced'}

        self.conf_er_code = { 0 : 'This device is working properly.', 1 : 'This device is not configured correctly.', 2 : 'Windows cannot load the driver for this device.', 
        3 : 'The driver for this device might be corrupted, or your system may be running low on memory or other resources.', 4 : 
        'This device is not working properly. One of its drivers or your registry might be corrupted.', 5 : 'The driver for this device needs a resource that Windows cannot manage.',
        6 : 'The boot configuration for this device conflicts with other devices.', 7 : 'Cannot filter.', 8 : 'The driver loader for the device is missing.',
        9 : 'This device is not working properly because the controlling firmware is reporting the resources for the device incorrectly.', 
        10 : 'This device cannot start.', 11 : 'This device failed.', 12 : 'This device cannot find enough free resources that it can use.', 
        13 : 'Windows cannot verify this device\'s resources.', 14 : 'This device cannot work properly until you restart your computer.', 
        15 : 'This device is not working properly because there is probably a re-enumeration problem.', 16 : 'Windows cannot identify all the resources this device uses.',
        17 : 'This device is asking for an unknown resource type.', 18 : 'Reinstall the drivers for this device.', 19 : 'Failure using the VxD loader.', 
        20 : 'Your registry might be corrupted.', 21 : 'System failure: Try changing the driver for this device. If that does not work, see your hardware documentation. Windows is removing this device.', 
        22 : 'This device is disabled.', 23 : 'System failure: Try changing the driver for this device. If that doesn\'t work, see your hardware documentation.', 
        24 : 'This device is not present, is not working properly, or does not have all its drivers installed.', 25 : 'Windows is still setting up this device.', 
        26 : 'Windows is still setting up this device.', 27 : 'This device does not have valid log configuration.', 28 : 'The drivers for this device are not installed.', 
        29 : 'This device is disabled because the firmware of the device did not give it the required resources.', 30 : 'This device is using an Interrupt Request (IRQ) resource that another device is using.',
        31 : 'This device is not working properly because Windows cannot load the drivers required for this device.' }

        self.prt_state = { 0 : 'Idle', 1 : 'Paused', 2 : 'Error', 3 : 'Pending Deletion', 4 : 'Paper Jam', 5 : 'Paper Out', 6 : 'Manual Feed', 
        7 : 'Paper Problem', 8 : 'Offline', 9 : 'I/O Active', 10 : 'Busy', 11 : 'Printing', 12 : 'Output Bin Full', 13 : 'Not Available', 14 : 'Waiting', 15 : 'Processing', 
        16 : 'Initialization', 17 : 'Warming Up', 18 : 'Toner Low', 19 : 'No Toner', 20 : 'Page Punt', 21 : 'User Intervention Required', 22 : 'Out of Memory', 23 : 'Door Open', 
        24 : 'Server_Unknown', 25 : 'Power Save' }
                
        self.hostname = hostname

        try:
            self.c = wmi.WMI(hostname)                             # Attempts to create a WMI session with remote computer
            objWMIService = win32com.client.Dispatch("WbemScripting.SWbemLocator")
            self.objSWbemServices = objWMIService.ConnectServer(hostname,"root\cimv2")
        except Exception as e:
            #print e
            print "Can't resolve the host or workstation is offline."
            x = raw_input("Press Enter to exit")
            exit(1)
        m = self.main()                                    # Calls the main function

    def last_build(self):
        """ This function is used to get the last build date on the remote workstation """
        print "---------------------------------"
        print "Basic information for workstation:"
        print "---------------------------------"
        file_list = []
        try:
            directory = "\\\\"+(self.hostname)+"File location"                  # Grabs the name of file in build_results folder
            for i in os.listdir(directory):
                a = os.stat(os.path.join(directory, i))
                file_list.append([time.ctime(a.st_atime)])
            print "Build Completed: " + str(file_list)
        except Exception as e:
            #print e
            print "Can not find build results."

    def os_system_info(self):
        """ This function is used to get basic operating system infomation for a remote workstation"""
        try: 
            for os in self.c.Win32_OperatingSystem():                               # Grabs operating system info from remote host
                os = os.Caption
            colItems = self.objSWbemServices.ExecQuery("SELECT * FROM Win32_ComputerSystem")
            for objItem in colItems:
                if objItem.Model != None:
                    pc_modle = str(objItem.Model)
            print os
            print pc_modle
            print self.hostname.upper()
        except Exception as e:
            #print e
            print "Unable to get operating system information."

    def system_uptime(self):
        """ This function is used to gather System Uptime information for a remote workstation"""
        try:
            secs_up = int([uptime.SystemUpTime for uptime in self.c.Win32_PerfFormattedData_PerfOS_System()][0])    # Grabs PC Uptime in seconds
            minutes_up = secs_up /60
            hours_up = secs_up / 3600                                  # Transfers the Uptime in seconds to days, hours, and minutes
            days_up = hours_up / 24
            if minutes_up > 0:
                day = days_up
                hour = hours_up - (day * 24)
                minute =  minutes_up - (hours_up * 60)
                print "Machine Uptime: ", day, "days", hour, "hours", minute, "minutes"
        except Exception as e:
            #print e
            print "Unable to pull uptime information."
        print

    def network_info(self):
        """ This function gathers network information for a remote workstation"""
        print "-------------------"
        print "Network Information:"
        print "-------------------"
        try:
            for interface in self.c.Win32_NetworkAdapterConfiguration (IPEnabled=1):  # Grabs and prints IP and MAC information
                print "MAC: " + interface.MACAddress
                print "IPv4: " + interface.IPAddress[0]
                #print "IPv6: " + interface.IPAddress[1]        
        except Exception as e:
            #print e
            print "Can not get " + self.hostname + "'s IP's or MAC Address."
            
        try:
            colItems = self.objSWbemServices.ExecQuery("SELECT * FROM Win32_NetworkAdapter")   # Grabs information from the netwok adaptor
            for objItem in colItems:
                if objItem.AdapterType != None:
                    x = str(objItem.AdapterType)
                    if x == 'Ethernet 802.3':
                        if objItem.Description != None:
                            print "NIC: " + str(objItem.Description)
                        if objItem.Availability != None:
                            num = int(objItem.Availability)
                            print "Availability: " + self.ava_code[num]
                        if objItem.ConfigManagerErrorCode != None:
                            num = int(objItem.ConfigManagerErrorCode)
                            print self.conf_er_code[num]
                        if objItem.Speed != None:
                            speed = str(objItem.Speed)
                            if speed == '1000000000':
                                print "Speed: 1GB"
                            elif objItem.Speed == '1000000':
                                print "Speed: 100MB"
                            else:
                                print "Speed: " + speed + " bps"
                        print 
        except Exception as e:
            #print e
            print "Unable to gather information from the network adapter."     

    def mem_cpu(self):
        """ This function is used to print Memory and CPU information for a remote workstation"""
        print "--------------------"
        print "Memory and CPU Usage:"
        print "--------------------"
        try:    
            for i in self.c.Win32_ComputerSystem():                      ### Try uptting all the Try statments in one
                totalMem = int(i.TotalPhysicalMemory)
                totalMem = totalMem / 1000000000
            pct_in_use = int([mem.PercentCommittedBytesInUse for mem in self.c.Win32_PerfFormattedData_PerfOS_Memory()][0])
            utilizations = [cpu.LoadPercentage for cpu in self.c.Win32_Processor()]
            utilization = int(sum(utilizations) / len(utilizations))  # avg all cores/processors
            print "Total memory: ", totalMem, "GB"
            print "Percent of free memory: ", pct_in_use, "%"
            print "Percent of CPU utilization: ", utilization, "%"
        except Exception as e:
            #print e
            print "Unable to pull memory and CPU information."
        print

    def motherboard_info(self):
        """ This function is used to gather motherboard information for a remote workstation"""
        print "-----------"
        print "Motherboard: "
        print "-----------"
        try:
            colItems = self.objSWbemServices.ExecQuery("SELECT * FROM Win32_BaseBoard")
            for objItem in colItems:
                if objItem.SerialNumber != None:
                    print "SerialNumber: " + str(objItem.SerialNumber)
                if objItem.Status != None:
                    print "Status: " + str(objItem.Status)
            colItems = self.objSWbemServices.ExecQuery("SELECT * FROM Win32_MotherboardDevice")
            for objItem in colItems:
                if objItem.Availability != None:
                    num = int(objItem.Availability)
                    print "Availability: " + self.ava_code[num]
        except Exception as e:
            #print e
            print "Unable to get Motherboard information."
        print

    def drive_size(self):
        """ This function is used to gather drive size for a remote workstation"""
        print "----------"
        print "Drive Size: "
        print "----------"
        try:
            for d in self.c.Win32_LogicalDisk():
                drive = str(d.Caption)
                if d.FreeSpace == None:
                    free = 0
                else:
                    free = float(d.FreeSpace)
                    free = float("{0:.2f}".format(free / 1073741824))
                if d.Size == None:
                    size = 0
                else:
                    size = int(d.size)
                    size = size / 1000000000
                print "Disk: ", drive, "     Free GB: ", free, "     Total GB: ", size
        except Exception as e:
            #print e
            print "Unable to get Drive Size information."
        print

    def drive_info(self):
        """ This function is used to gather drive information for a remote workstation"""
        print "-----------"
        print "Disk Drives: "
        print "-----------"
        try:
            colItems = self.objSWbemServices.ExecQuery("SELECT * FROM Win32_DiskDrive")
            for objItem in colItems:
                if objItem.Model != None:
                    print "Model: " + str(objItem.Model)
                if objItem.Name != None:
                    print "Name: " + str(objItem.Name)
                if objItem.InterfaceType != None:
                    print "InterfaceType: " + str(objItem.InterfaceType)
                if objItem.Status != None:
                    print "Status: " + str(objItem.Status)
                if objItem.ConfigManagerErrorCode != None:
                    num = int(objItem.ConfigManagerErrorCode)
                    print self.conf_er_code[num]
                if objItem.Availability != None:
                    num = int(objItem.Availability)
                    print "Availability: " + self.ava_code[num]
                print
        except Exception as e:
            #print e
            print "Unable to get disk drive information."
        print

    def fan_info(self):
        """ This function is used to gather fan information for a remote workstation"""
        print "---------------"
        print "Fan information: "
        print "---------------"
        try:
            colItems = self.objSWbemServices.ExecQuery("SELECT * FROM Win32_Fan")
            for objItem in colItems:
                if objItem.Name != None:
                    print "Name: " + str(objItem.Name)
                if objItem.Availability != None:
                    num = int(objItem.Availability)
                    print "Availability: " + self.ava_code[num]
                if objItem.Status != None:
                    print "Status: " + str(objItem.Status)
                if objItem.ConfigManagerErrorCode != None:
                    num = int(objItem.ConfigManagerErrorCode)
                    print self.conf_er_code[num]
        except Exception as e:
            #print e
            print "Unable to get Fan information."
        print

    def usb_info(self):
        """ This function is used to gather USB information for a remote workstation"""
        print "---------------------"
        print "USB ports on computer:"
        print "---------------------"
        try:
            for usb in self.c.InstancesOf("Win32_UsbHub"):      # Runs a loop to get the names of each USB on the remote host
                print 'Name: ' + usb.Name
        except Exception as e:
            #print e
            print "Unable to gather USB information."
        print

    def printer_info(self):
        """This funtion is used to gather printer information from a remote workstation"""
        colItems = self.objSWbemServices.ExecQuery("SELECT * FROM Win32_Printer")
        print "-------------------"
        print "Printer Information:"
        print "-------------------"
        print
        try:
            for objItem in colItems:
                strList = " "
                try:
                    print "CapabilityDescriptions: ",
                    for objElem in objItem.CapabilityDescriptions :
                        strList = strList + str(objElem) + ", "
                except:
                    strList = strList + 'null'
                print strList
                if objItem.Caption != None:
                    print "Caption: " + str(objItem.Caption)
                if objItem.DriverName != None:
                    print "DriverName: " + str(objItem.DriverName)
                if objItem.ConfigManagerErrorCode != None:
                    num = int(objItem.ConfigManagerErrorCode)
                    print "Error Status: " + self.conf_er_code[num]   
                if objItem.PrinterState != None:
                    num = int(objItem.PrinterState)
                    print "Print State: " + self.prt_state[num]
                if objItem.Shared != None:
                    print "Shared: " + str(objItem.Shared)
                if objItem.ShareName != None:
                    print "ShareName: " + str(objItem.ShareName)
                if objItem.SpoolEnabled != None:
                    print "SpoolEnabled: " + str(objItem.SpoolEnabled)
                print
        except Exception as e:
            #print e
            print "Unable to get printer information"

    def socket_connect(self):
        """ This function is used to attempt a socket creation for a remote workstation"""
        print "-----------------"
        print "Socket Connection:"
        print "-----------------"
        print
        try: 
            s = socket.socket(socket.AF_INET, socket.SOCK_STREAM) # Calls socket
            s.connect((self.hostname, 3389))                           # Connects to socket on port 3389

            print "Socket connection successfully, should be able to RDP." 
            s.close() 
        except Exception as e:
            #print e 
            print "Socket failed to connect." 
        print

    def ping_results(self):
        """ This function is used to attempt a ping to a remote workstation"""
        print "----"
        print "Ping:"
        print "----"
        try:
            socket.gethostbyname(self.hostname)                # Uses socket to resolve host 
            output = check_output(["ping", self.hostname])     # Runs ping for Windows
            print output
        except Exception as e:
            #print e
            print "Unable to Ping."

    def main(self):
        """ Runs all functions to gather the data for a remote workstation"""
        self.last_build()
        self.os_system_info()
        self.system_uptime()
        self.network_info()
        self.mem_cpu()
        self.motherboard_info()
        self.drive_size()
        self.drive_info()
        self.fan_info()
        self.usb_info()
        self.printer_info()
        self.socket_connect()
        self.ping_results()
        x = raw_input("Press Enter to exit")

os.system('cls')                                         # Clears the screen
hostname = raw_input('Enter the computer hostname: ')    # Provides input for hostname to gather information
pc = PcHealth(hostname)                                          # Calls The class 
