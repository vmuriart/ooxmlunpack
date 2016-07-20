# Storing Excel Files Efficiently #

http://hgbook.red-bean.com/read/behind-the-scenes.html#x8-640004 - provides information about how Mercurial stores revisions of files internally.

A file that is large, or has a lot of history, has its filelog stored in separate data (“.d” suffix) and index (“.i” suffix) files. For small files without much history, the revision data and index are combined in a single “.i” file.

So, if we look at the size of the .i and .d files corresponding to a recently committed change, before and after, we can infer the (binary) size of the commit.

In particular, I'm interested here in how to efficiently store revisions of spreadsheets, which are (in the latest Open Document Format) stored as .xlsx or .xlsm files, which are zip archives containing mostly XML files. In fact, if you take an .xlsx file and rename it as .zip, you can open or uncompress it using regular zip tools.

It does appear that older (.xls) Excel files actually compress quite well in Mercurial - the Template Manager file E.Swing.xls at time of writing has 7 revisions of changes to a file that is 3.92MB, all stored in just 3.27MB:

This, I think, is because the xls is a raw binary format, so Mercurial is able to diff it quite effectively. On the other hand, since the xlsx format is zipped, small changes in the raw XML could mean big changes in the zipped file.

Good news though, is that if you decompress an xlsx file and then create a new Zip archive from the contents but with no compression, you get a file that Excel will still open, albeit a much larger file. Saving the 4MB E.Swing.xls file as E-Swing.xlsm (it has macros) gives a new file of less than 2MB, however applying this process to create a new .xlsm file that is zipped but not compressed, gives a file of 21MB which Excel will open. Open the original .xls or .xlsm files into Notepad++ and you see binary garbage, but the expanded (and still valid) .xlsm contains a lot of recognisable plain text.

http://msdn.microsoft.com/en-us/library/office/gg278309%28v=office.15%29.aspx - provides details of the XML format that Excel uses internally to represent a spreadsheet. This is an open standard, so ought to be able to be relied upon.

Now, there already exists an extension for Mercurial called ZipDoc (https://bitbucket.org/gobell/hg-zipdoc) which hooks into the [encode] and [decode] events in Mercurial to allow specific file types to be decompressed before commit and recompressed on update. Someone has blogged about it not working very well for Excel files though:

http://www.devuxer.com/2014/02/why-the-mercurial-zipdoc-extension-fails-for-excel-files/#.U422SyhNoiU

His point is that the XML files within the .xlsx archive are not pretty printed, so there are no line breaks and therefore the internall diff encoding of the revset fails to calculate the diff correctly. I have to say I'm not convinced that the problem is as he describes, given the good compression Mercurial seems to be able to achieve on .xls binary files, but I'm also not totally convinced that ZipDoc is working as advertised. I also see the benefit of pretty-printing the XML since it would nice to be able to visualise the diff, as well as have Mercurial represent the revision history without it taking up loads of space.

So, I've build a small .NET executable (Mercurial expects extensions to be written in python, but I'm not a python expert!) which will transform an .xlsx file into a new, larger file which is a valid file that Excel can open, but which has had its compression removed, and every file within the archive which can be parsed successfully as XML has been loaded as an XDocument and then saved again, which has the side-effect of pretty-printing the XML. Applying this to my 21MB E.Swing.xlsm file and checking it in a few times, I now have five revisions of my 21MB file stored in 1.9MB:

Also, I can get useful side-by-side diffs from one revision to the next (okay, they're not perfect but a massive step forward):

One other thing I've done as part of this is to remove the file calcChain.xml from the zip archive. This is a file that describes the calculation order in the file and is optional. It can change a lot just from recalculating the workbook so it introduces a lot of spurious differences between one version and the next. Without it, the sheet still opens fine into Excel, so it seems sensible to discard this file as part of the encoding process.