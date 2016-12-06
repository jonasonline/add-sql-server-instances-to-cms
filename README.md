# add-sql-server-instances-to-cms
A script for adding instances of Microsoft SQL Server to a [Central Management Server](https://msdn.microsoft.com/en-us/library/bb934126.aspx).

## Usage
Run the script once to generate a configurtion file.
Edit the configuration file `config.json` with the parameters for your environment.

### Configuration Parameters
* **UncategorizedServerGroup** Server group in the CMS for newly added servers.
* **InventoryFilePath** Path to csv files containing instance information. The csv files have to contain tree columns for each instance. 'MachineName', 'DisplayName' _(ex. SQL Server (MSSQLSERVER))_ and 'Status' _(Running/Stopped)_.
* **ErrorLogPath** Path to the error log for the script.
* **CMS** Name of the SQL Server instance acting as CMS.  