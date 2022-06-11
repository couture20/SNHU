## Welcome to Ryan Couture's ePortfolio

The purpose of this page is to summarize and briefly showcase my abilities as a programmer. Below you'll see examples of past and ongoing projects that I have done over the past few years. They're examples of projects that completed for my current job, of applications done while in college, along with projects that I do in my spare time at home. 

For the past few years I have been expanding my knowledge of programming, through college work, my current job at Ups and through home projects such as NFTs. 

Currently I work at Ups as the Central Zone I.E. Service Engineer. My current duties are but not limited to:
- Data mining via either code that I have written or through advance interactions with Oracle and Power Bi. I am very fluent in SQL programming, and can generate quries through numerous different interfaces. 
- I create custom reports for Corporate management, that often require coding in VBA that interact with either hardlined or cloud computing databases.
- Evaluate new technologies to be implemented within the zone. Whether it's new automation systems or new applications, I evaluate if it can enhance the current systems and what the potential ROI would be on the new system or application.
- Because of my interactions with emerging automation systems, I am able to implement and troubleshoot XLE and PLC systems.
- Through these various aspects of my job, I frequently work within a team or lead a team depending on the project. 
- I train and develop operations managment teams and other company employees on how to effectively use newly implememted systems.
- Ovrall my role is to monitor and enhance the zones logisitc capabilities and to ensure a more profitable and quick return on investments. 

In my personal life I have found in interest in blockchaining, in particular with NFTs. I have made several basic websites that utilize personalized code to allow the website's visitors to mint NFTs that I created and send it to their digital wallet. I have written code that will generate thousands of various NFTs based on the users layer inputs, and will attach the metadata to each NFT. Below is a link to a NFT minting app that I created, along with the artwork that was generated through the art generating program.

[Cthulhu Has Risin NFT](https://cthulhuhasrisin.com/)




### VBA code for report generating

Below is a brief example snippet of VBA code that I created that interacts with the comapnies SQL Database. The purpose of it, was to parse and format a users data input for a SQL pass through query and then output the data in a user friendly readable table. I created it roughly a year ago with the purpose of minimizing labor intesive data collection through ineffcient company web applications. After creatin gthis report I was able to reduce several hours of dataminning and report generating, down to just a few minutes. I wanted to include this bit of code because it showcases my ability to design and engineer a program that interacts with a database through a program language (VBA) that is utilized through the widely used Microsoft 365 suite. For the purposes of company proprietary information, pieces of the code have been either altered or removed in order to protect comapny data.

        Public Sub GssData()

        'First set of variables define and interact with the SQL Database
        Dim CoNStr, StrSQL As String
        Dim ConSvr As ADODB.Connection
        Dim RSSVR As ADODB.Recordset

        'Declared variables for worksheets names
        Dim sh1, sh2 As Worksheet

        'Declared variables to find last use row and last used column
        Dim lastCol, lastRow As Long

        'Declared variables for the while loop counter
        Dim x As Long

        'Decloration for variables used for SQL query
        Dim trk As Range
        Dim origin, sort As String

        'Sets variables for a new connection to a database and for a new record
        Set ConSvr = New ADODB.Connection
        Set RSSVR = New ADODB.Recordset

        'Defines the variables, that way we only have to use sh1 or sh2 to reference different sheets
        Set sh1 = ThisWorkbook.Sheets("String")
        Set sh2 = ThisWorkbook.Sheets("GSS")

        'Finds the last used row and clears the contents
        lastRow = sh2.Cells(Rows.Count, 1).End(xlUp).Row
        sh2.Range("A2:H" & lastRow).Clear

        'Finds the last used column
        With sh1
            lastCol = .Cells(1, .Columns.Count).End(xlToLeft).Column
        End With

        'Defines the variables that will be used in the SQL query
        origin = sh1.Range("A3")
        sort = sh1.Range("A4")

        'This will keep the connection open to the Databse for 120 seconds
        ConSvr.CommandTimeout = 120

        'An error handler for if the user enters the wrong usernam and/or password
        On Error GoTo Broke
        'Passes the connection string and opens the connection up to the Database
        CoNStr = "DSN=GSS_Arch;UID=UsersID;PWD=UsersPwd;DATABASE=GSSDataArchive;"
        ConSvr.Open CoNStr

        'Sets the variable that will be used as a counter for the while loop below
        x = 1

        'A while loop that will loop through each column of data until the last column of data is used.
             Do While x <= lastCol
                'will progress as the while loop does, ensuring that each row and column of data is processed through the query.
                Set trkFind = sh1.Range(Cells(x, lastCol), Cells(x, lastCol))
                'The string that is passed through to the SQL database in order to run the defined query.
                StrSQL = "SELECT PKG_TCK_NR, SRT_DT, SRT_TYP_CD, ORG_OGZ_NR, DTN_OGZ_NR,DTN_SRT_TYP_CD,BLT_NR,BAY_NR " & vbCrLf & _
                        "FROM dbo.VSCN_PSL " & vbCrLf & _
                        "WHERE PKG_TCK_NR IN ('" & trkFind & " 'null') AND ORG_OGZ_NR='" & origin & "' AND SRT_TYP_CD ='0" & sort & "'"
                RSSVR.Open StrSQL, ConSvr
                'Pastes the record set to sh2
                sh2.Range("A1:A" & lastRow).CopyFromRecordset RSSVR
                RSSVR.Close
                'counts for the while loop
                x = x + 1
            Loop

        'Closes out the recordeset and the string
        Set RSSVR = Nothing
        Set ConSvr = Nothing

        'Adds headers in for the data pulled from the Database. It's used because it's more user friendly than the DB headers.
        sh2.Range("A1") = "Tracking"
        sh2.Range("B1") = "Date"
        sh2.Range("C1") = "Origin Sort"
        sh2.Range("D1") = "Origin"
        sh2.Range("E1") = "Dest"
        sh2.Range("F1") = "Dest Sort"
        sh2.Range("G1") = "Area"
        sh2.Range("H1") = "Bay"


        'Redefines the variable to find a new last row
        lastRow = sh2.Cells(Rows.Count, 1).End(xlUp).Row
        'Clears out column 9
        sh2.Columns(9).EntireColumn.Clear
        'Writes an excel formula that parses the Data from the Database and autfills it down the column
        sh2.Range("I2").Formula = "=If(left(G2,2)=""PD"",left(G2,2)&right(G2,2),G2)"
        sh2.Range("I2").AutoFill Destination:=sh2.Range("I2:I" & lastRow)
        'Copies the formulas output and pastes just the values in another column
        sh2.Range("I2:I" & lastRow).Copy
        sh2.Range("G2").PasteSpecial xlPasteValues
        Application.CutCopyMode = False

        Exit Sub

        'Error Handler for if the user enters the wrong username or password
        Broke:
                MsgBox "Invalid Username or Password"

        End Sub




### Simple Password Protect

Below is another brief example snippet of a C++ code that I made for a College class in my Computer Science degree program. In it, I made a very simple yet effective while loop that asks the user to input the password before they are able to use the application. The user will have 3 attempts before the program will end. The snippet is part of a larger program that allows the user to interact with product bids. I included this bit of code because showcases the applications use of a variety of alorithms and data structures by utilizing nodes, nested loops, read/writes to a .csv and user interactions.

    cout << "===== Welcome! Please enter the Password! =====" << endl;
    cout << "===== You only get 3 Attempts! =====" << endl;

    //A while statment to allow the user 3 attempts for the password
    While (attCount < 3 ){
        cout << "Password: " << endl;
        cin >> pwd;

        //Sets the password and is an if statment to allow the user to input password
        if (pwd !="12345"){
            cout << "Incorrect Password, Please Try Again!" << endl;
            attCount++;
        }
        else{

            cout << "Correct Password!" << endl;
            system("cls")
        }
    }
    //If the password is true then it will continue to the while loop
    If (pwd == "12345"){
        cout << "Welcome! Please select your choice." << endl;
        int choice = 0;

        //The while loop present the user with options
        while (choice != 9) {
            cout << "Menu:" << endl;
            cout << "  1. Load Bids" << endl;
            cout << "  2. Display All Bids" << endl;
            cout << "  3. Find Bid" << endl;
            cout << "  4. Remove Bid" << endl;
            cout << "  9. Exit" << endl;
            cout << "Enter choice: ";
            cin >> choice;

            //Switch statment to run through the user options
            switch (choice) {

            case 1:
                bst = new BinarySearchTree();

                // Initialize a timer variable before loading bids
                ticks = clock();

                // Complete the method call to load the bids
                loadBids(csvPath, bst);

                //cout << bst->Size() << " bids read" << endl;

                // Calculate elapsed time and display result
                ticks = clock() - ticks; // current clock ticks minus starting clock ticks
                cout << "time: " << ticks << " clock ticks" << endl;
                cout << "time: " << ticks * 1.0 / CLOCKS_PER_SEC << " seconds" << endl;
                break;

            //Puts bids in order
            case 2:
                bst->InOrder();
                break;

            //Finds the bids
            case 3:
                ticks = clock();

                bid = bst->Search(bidKey);

                ticks = clock() - ticks; // current clock ticks minus starting clock ticks

                if (!bid.bidId.empty()) {
                    displayBid(bid);
                } else {
                    cout << "Bid Id " << bidKey << " not found." << endl;
                }

                cout << "time: " << ticks << " clock ticks" << endl;
                cout << "time: " << ticks * 1.0 / CLOCKS_PER_SEC << " seconds" << endl;

                break;

            //Removes the bid
            case 4:
                bst->Remove(bidKey);
                break;
            }
        }
    }

    //Exits the program if too many attempts were made
    else{
        system("cls");
        cout << "You have reached you 3 attempts. The program will exit now" << endl;
    }
    cout << "Good bye." << endl;

	return 0;
        
        
```markdown
Syntax highlighted code block

# Header 1
## Header 2
### Header 3

- Bulleted
- List

1. Numbered
2. List

**Bold** and _Italic_ and `Code` text

[Link](url) and ![Image](src)
```

For more details see [Basic writing and formatting syntax](https://docs.github.com/en/github/writing-on-github/getting-started-with-writing-and-formatting-on-github/basic-writing-and-formatting-syntax).

### Jekyll Themes

Your Pages site will use the layout and styles from the Jekyll theme you have selected in your [repository settings](https://github.com/couture20/SNHU/settings/pages). The name of this theme is saved in the Jekyll `_config.yml` configuration file.

### Support or Contact

Having trouble with Pages? Check out our [documentation](https://docs.github.com/categories/github-pages-basics/) or [contact support](https://support.github.com/contact) and we’ll help you sort it out.
