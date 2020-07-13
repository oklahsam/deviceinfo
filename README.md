# deviceinfo
Script to show info on AD computers, and show what switch ports they are plugged into


This one has taken me a while to get it to the point where I felt I could share it.

It requires some extra modules, but it checks for and installs them when you run it, so as long as you run the script as an admin you should be alright there.

For it to work properly, it has to be able to find the MAC address of the device you're looking at. It looks at DHCP on the domain controller first. If it doesn't find it there, it looks through the switches using SNMP. This can take quite a bit longer, especially if you've got a lot of Cisco switches with multiple VLANs since the script has to look through each VLAN separately by using "SNMPcommunity@VLAN" (a nightmare on our 4x stack of 3850s with about 10 VLANS). This isn't an issue on the Aruba switches I tested.

There is also a second tab where you can search by IP or MAC address manually, so you can look for devices outside of just AD computers.

I have tested this on Cisco 3750/3850s and various Aruba switches and it works great. I did also test it on my Ubiquiti switch at home and it didn't work there, so YMMV.


A couple more notes:
If the ports on the switch have a description, it will display it.

You can click on the MAC address it finds for the device to copy it to your clipboard.
