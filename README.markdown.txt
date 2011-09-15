# Firewall Updater service

This project contains a service that will at regular intervals perform a lookup on a host (typically a dynamic DNS service like NoIP.org or DynDNS) to discover the IP, and update windows advanced firewall using powershell to allow a configured port through.

I use this service on a remote server to update the firewall to allow my home broadband IP address to use Windows Remote Desktop to connect.