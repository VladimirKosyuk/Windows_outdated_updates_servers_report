# Report_3_month_old_updated_windows_servers

Should be usefull th know if WSUS administator do not update all production servers during 3 month period:

•	Forms a list of those servers whose objects are turned on, accessed DC no more than 14 days ago, whose OU contains servers and not contains test;

•	For each of received servers via WinRM script gets a list of updates;

•	If no values are received for server during 5 minutes - report to log file and continue with another server;
•	If errors were received during data collection - output to the log file;

•	If server received updates older than 3 months - write the information about the latest update to the csv file;

•	If server received updates no older than 3 months - continue to another server;

•	Send an email without authorization with results.
