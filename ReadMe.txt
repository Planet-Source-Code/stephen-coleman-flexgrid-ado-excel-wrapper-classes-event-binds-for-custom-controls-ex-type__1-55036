Wrapper Classes for Common Controls - ADODB and MSFLEXGRID
This project is designed for intermediate users. It assumes you under stand OOP.
I know I’ve seen many examples of the functions in the classes I’ve included. I didn't see an example of how to handle all of the usually required functions of a flexgrid and SQL\MDB connection. This project contains the functions I always need when using flexgrids and ADODB in a user interface application. This might help someone who is wondering what is the difference between a class and a module and which one do I use. Assuming you know how to use a module this example will show you how to use a class.
Project Requires Excel, and MDAC look at the refrences.bmp to see them. Tested with MDAC 2.6,2.7,2.8 and Excel 97,2003
Access database used is "C:\Program Files\Microsoft Visual Studio\VB98\NWIND.MDB" You can easily make an MDB and populate data. Then change the query or point the C_CQL to a SQL server.database.

Listed are some of the positives and negatives of classes.
Positives:
1.You can have many instances of a class. 
a. Ex C_SQL class you could Dim 2 instances of it as CSQL1 and CSQL2. Then connect CSQL1 to server ABC andCSQL1 to Server XYZ. Now using both instances of the classes you could retrieve data from both servers and manipulate it.
2. Wrapper classes can make Application.exe size smaller. You have 30 forms with a flexgrid control. Each from has a flexgrid using a sorting function, in a mouse click event on each from. One wrapper class saves you 30 on mouse click events.
2. You can easily kill the class by setting it to nothing and retrieve all the used memory.
3. Faster Development time with Reusable non-cut and paste code. With well-written classes you can make applications faster and more stable. Never having to cut and paste a function from an old app or a code bank, then make changes to it saves a lot of time, and eliminates errors.

Negatives:
1. Having two classes open takes up more memory than just creating two recordsets and retrieving the data from one connection.
2. Application.exe size can be larger because of unused functions in the class. However removing the unused functions could eliminate this problem.

Project covers:
WithEvents to hook into the events of a control.
	Dynamic adding of controls to a form
Advanced Flexgrid features;
	Outputting Flexgrid  to excel
	Outputting Flexgrid  to File
	creating and using check boxes on a flexgrid "Credit to someone can't remember who"
Advanced database recordset usage
Transferring Data from database to Flexgrid using Clip controls
Advanced Excel usage
Using Excel to manipulate spreadsheets
	Coping Flexgrid data into excel and formatting it.
Advanced Error logging
	Email notification on errors
	Using the C_Error Class in conjunction with the top of procedure error trap makes adding error handling a breeze. I make it even 

easier using : "Code Inserter ToolWindow" Ver 1.00 By Dean J. Giovanelli
	This is my Code Inserter macro for module errs, I don't even have to type in the procedure or module name.
	{<PosFirstLineProc>
	On Error GoTo Error_Handler '**Error Trap**
	If Err.Number Then
	Error_Handler:
	     ErrC.errortrap "<ModuleName>", "<ProcedureName>", Err.Number, Err.Description, Err.Source
	     Resume Next
	End If '**Error Trap**}
Click the top rows:  The Flex grid has a A -> Z and Z -> A sort 
Select a column and Type ahead just like in Excel.
	

