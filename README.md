# PreventativeMaintenanceManager_03-10-2014

Date Written:03/10/2014

Industry: Time of Flight Mass Spectrometer Developer & Manufacturer

Department: Hardware & Software Customer Support

GUI: Little to no energy was put forth in developing a robust GUI because I was the end user.

Sample Raw Data: A spreadsheet maintained by the departmental secretary.

Sample Output:

Emails were sent to each of the appropriate service engineers listing their assigned PM customer site visits with useful Metadata (“SampleEmail_Engineer1.png”, “SampleEmail_Engineer2.png”, & “SampleEmail_Engineer3.png”), and a master email to the manager listing all the assigned PM’s for each service engineer (“SampleEmail_ManagerSummary.png”).  A database maintained a record of all assigned tasks.  This image (“ServiceEngineerPM-Table.png”) showed a query for a single service engineer for about 1 year of data. 

Application Description:

When executed, this application would open and check a spreadsheet for all customers who were owed a Preventative Maintenance (PM) site visit.  Of the customers that were, their Metadata was mined, and appended to a data structure.  Next, the appropriate service engineer was matched to a given customer based on the engineer’s availability and current location.  Then emails would be drafted and set to each service engineer listing their PM assignments including useful customer Metadata.  The service department manager would also be sent an email with a master list of each service engineer’s PM assignments.  All this information was stored in a database for tracking and accountability.

The biggest challenge with this particular project was that I was not given the freedom to redesign the spreadsheet that contained all the customer PM data.  That spreadsheet existed for years prior to the development of this application, and was maintained by our departmental secretary.  The organizational structure for tracking information was loosely laid out and didn’t not follow a straight forward method for defining which customer was owed a PM.  As such the algorithm necessary to mine the correct customer information was by far the most difficult part of this project.

