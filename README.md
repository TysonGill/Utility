# Utility

So after perhaps 200+ casual projects and 100+ major commercial projects, this homebrew has been my secret weapon since the early 1990's. Of course then it was just a cut and paste code library, bu regardless, it has always enabled me to do very reliable projects quickly with a high degree of reuse and therefore a small footprint of code needing testing. It contains a wide away of methods. Some are things I use all the time in every project, and others are things I do rarely enough that I would forget how to do them and rather not have to relearn. Some things have since dropped out when they were superceded by enhancements to the native language but many remain very useful.

As just one example, the DataQuery object makes fast and robust database programming trivially simple with extremely little code:

Dim dq as clsDataQuery = New clsDataQuery(ConnectionString) ' do this once at the start of your program to initialize

Then, whenever you need...

Dim dt as DataTable = dq.GetTable(sql) ' Return a datatable

obj = dq.GetValue(sql) ' Return a single field or aggregate value

int = dq.GetIdentity(sql) ' Insert a new record and return the identity

dq.Execute(sql) ' Execute an action query

etc

There are similar libraries for database and datatable manipulation, file system management, formatting, mapping, multimedia, networking, SQL parsing, and more. Most of the other projects in my repository rely upon it heavily. I rarely touch it, but that's the good thing. I rarely need to.


![image](https://user-images.githubusercontent.com/120231132/206852471-81340ab1-4e3c-4903-9d7d-391187bb58ae.png)
