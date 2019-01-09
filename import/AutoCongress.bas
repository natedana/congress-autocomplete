Attribute VB_Name = "AutoCongress"

Sub AutoCongress()
Attribute AutoCongress.VB_Description = "Auto Complete for 2018 Feb. House and Senate"
Attribute AutoCongress.VB_ProcData.VB_Invoke_Func = "Nates_Templates.NewMacros.AutoCongress"
'
' AutoCongress Macro
' Auto Complete for 2018 Feb. House and Senate
' START: Dim oAutoText As AutoTextEntry
' ENTRY: Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
'       .Add(Name:="NAME", Range:=Selection.Range)
'   oAutoText.Value = "LONG FORM"
' END: Set oAutoText = Nothing

Call AutoSenate
Call AutoRep1
Call AutoRep2
Call AutoRep3

End Sub

Private Sub AutoSenate()
Dim oAutoText As AutoTextEntry

Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Alexander, Lamar", Range:=Selection.Range)
    oAutoText.Value = "Senator Lamar Alexander (R-TN)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Ayotte, Kelly", Range:=Selection.Range)
    oAutoText.Value = "Senator Kelly Ayotte (R-NH)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Baldwin, Tammy", Range:=Selection.Range)
    oAutoText.Value = "Senator Tammy Baldwin (D-WI)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Barrasso, John", Range:=Selection.Range)
    oAutoText.Value = "Senator John Barrasso (R-WY)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Baucus, Max", Range:=Selection.Range)
    oAutoText.Value = "Senator Max Baucus (D-MT)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Begich, Mark", Range:=Selection.Range)
    oAutoText.Value = "Senator Mark Begich (D-AK)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Bennet, Michael", Range:=Selection.Range)
    oAutoText.Value = "Senator Michael Bennet (D-CO)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Blumenthal, Richard", Range:=Selection.Range)
    oAutoText.Value = "Senator Richard Blumenthal (D-CT)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Blunt, Roy", Range:=Selection.Range)
    oAutoText.Value = "Senator Roy Blunt (R-MO)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Boozman, John", Range:=Selection.Range)
    oAutoText.Value = "Senator John Boozman (R-AR)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Boxer, Barbara", Range:=Selection.Range)
    oAutoText.Value = "Senator Barbara Boxer (D-CA)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Brown, Sherrod", Range:=Selection.Range)
    oAutoText.Value = "Senator Sherrod Brown (D-OH)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Burr, Richard", Range:=Selection.Range)
    oAutoText.Value = "Senator Richard Burr (R-NC)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Cantwell, Maria", Range:=Selection.Range)
    oAutoText.Value = "Senator Maria Cantwell (D-WA)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Cardin, Ben", Range:=Selection.Range)
    oAutoText.Value = "Senator Ben Cardin (D-MD)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Carper, Tom", Range:=Selection.Range)
    oAutoText.Value = "Senator Tom Carper (D-DE)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Casey, Bob", Range:=Selection.Range)
    oAutoText.Value = "Senator Bob Casey (D-PA)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Chambliss, Saxby", Range:=Selection.Range)
    oAutoText.Value = "Senator Saxby Chambliss (R-GA)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Chiesa, Jeff", Range:=Selection.Range)
    oAutoText.Value = "Senator Jeff Chiesa (R-NJ)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Coats, Dan", Range:=Selection.Range)
    oAutoText.Value = "Senator Dan Coats (R-IN)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Coburn, Tom", Range:=Selection.Range)
    oAutoText.Value = "Senator Tom Coburn (R-OK)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Cochran, Thad", Range:=Selection.Range)
    oAutoText.Value = "Senator Thad Cochran (R-MS)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Collins, Susan", Range:=Selection.Range)
    oAutoText.Value = "Senator Susan Collins (R-ME)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Coons, Chris", Range:=Selection.Range)
    oAutoText.Value = "Senator Chris Coons (D-DE)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Corker, Bob", Range:=Selection.Range)
    oAutoText.Value = "Senator Bob Corker (R-TN)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Cornyn, John", Range:=Selection.Range)
    oAutoText.Value = "Senator John Cornyn (R-TX)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Crapo, Michael", Range:=Selection.Range)
    oAutoText.Value = "Senator Michael Crapo (R-ID)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Cruz, Ted", Range:=Selection.Range)
    oAutoText.Value = "Senator Ted Cruz (R-TX)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Donnelly, Joe", Range:=Selection.Range)
    oAutoText.Value = "Senator Joe Donnelly (D-IN)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Durbin, Richard", Range:=Selection.Range)
    oAutoText.Value = "Senator Richard Durbin (D-IL)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Enzi, Michael", Range:=Selection.Range)
    oAutoText.Value = "Senator Michael Enzi (R-WY)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Feinstein, Dianne", Range:=Selection.Range)
    oAutoText.Value = "Senator Dianne Feinstein (D-CA)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Fischer, Deb", Range:=Selection.Range)
    oAutoText.Value = "Senator Deb Fischer (R-NE)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Flake, Jeff", Range:=Selection.Range)
    oAutoText.Value = "Senator Jeff Flake (R-AZ)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Franken, Al", Range:=Selection.Range)
    oAutoText.Value = "Senator Al Franken (D-MN)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Gillibrand, Kirsten", Range:=Selection.Range)
    oAutoText.Value = "Senator Kirsten Gillibrand (D-NY)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Graham, Lindsey", Range:=Selection.Range)
    oAutoText.Value = "Senator Lindsey Graham (R-SC)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Grassley, Chuck", Range:=Selection.Range)
    oAutoText.Value = "Senator Chuck Grassley (R-IA)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Hagan, Kay", Range:=Selection.Range)
    oAutoText.Value = "Senator Kay Hagan (D-NC)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Harkin, Tom", Range:=Selection.Range)
    oAutoText.Value = "Senator Tom Harkin (D-IA)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Hatch, Orrin", Range:=Selection.Range)
    oAutoText.Value = "Senator Orrin Hatch (R-UT)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Heinrich, Martin", Range:=Selection.Range)
    oAutoText.Value = "Senator Martin Heinrich (D-NM)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Heitkamp, Heidi", Range:=Selection.Range)
    oAutoText.Value = "Senator Heidi Heitkamp (D-ND)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Heller, Dean", Range:=Selection.Range)
    oAutoText.Value = "Senator Dean Heller (R-NV)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Hirono, Mazie", Range:=Selection.Range)
    oAutoText.Value = "Senator Mazie Hirono (D-HI)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Hoeven, John", Range:=Selection.Range)
    oAutoText.Value = "Senator John Hoeven (R-ND)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Inhofe, James", Range:=Selection.Range)
    oAutoText.Value = "Senator James Inhofe (R-OK)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Isakson, Johnny", Range:=Selection.Range)
    oAutoText.Value = "Senator Johnny Isakson (R-GA)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Johanns, Mike", Range:=Selection.Range)
    oAutoText.Value = "Senator Mike Johanns (R-NE)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Johnson, Ron", Range:=Selection.Range)
    oAutoText.Value = "Senator Ron Johnson (R-WI)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Johnson, Tim", Range:=Selection.Range)
    oAutoText.Value = "Senator Tim Johnson (D-SD)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Kaine, Tim", Range:=Selection.Range)
    oAutoText.Value = "Senator Tim Kaine (D-VA)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="King, Angus", Range:=Selection.Range)
    oAutoText.Value = "Senator Angus King (I-ME)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Kirk, Mark Steven", Range:=Selection.Range)
    oAutoText.Value = "Senator Mark Steven Kirk (R-IL)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Klobuchar, Amy", Range:=Selection.Range)
    oAutoText.Value = "Senator Amy Klobuchar (D-MN)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Landrieu, Mary", Range:=Selection.Range)
    oAutoText.Value = "Senator Mary Landrieu (D-LA)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Leahy, Pat", Range:=Selection.Range)
    oAutoText.Value = "Senator Pat Leahy (D-VT)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Lee, Mike", Range:=Selection.Range)
    oAutoText.Value = "Senator Mike Lee (R-UT)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Levin, Carl", Range:=Selection.Range)
    oAutoText.Value = "Senator Carl Levin (D-MI)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Manchin, Joe", Range:=Selection.Range)
    oAutoText.Value = "Senator Joe Manchin (D-WV)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Markey, Ed", Range:=Selection.Range)
    oAutoText.Value = "Senator Ed Markey (D-MA)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="McCain, John", Range:=Selection.Range)
    oAutoText.Value = "Senator John McCain (R-AZ)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="McCaskill, Claire", Range:=Selection.Range)
    oAutoText.Value = "Senator Claire McCaskill (D-MO)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="McConnell, Mitch", Range:=Selection.Range)
    oAutoText.Value = "Senator Mitch McConnell (R-KY)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Menendez, Robert", Range:=Selection.Range)
    oAutoText.Value = "Senator Robert Menendez (D-NJ)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Merkley, Jeff", Range:=Selection.Range)
    oAutoText.Value = "Senator Jeff Merkley (D-OR)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Mikulski, Barbara", Range:=Selection.Range)
    oAutoText.Value = "Senator Barbara Mikulski (D-MD)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Moran, Jerry", Range:=Selection.Range)
    oAutoText.Value = "Senator Jerry Moran (R-KS)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Murkowski, Lisa", Range:=Selection.Range)
    oAutoText.Value = "Senator Lisa Murkowski (R-AK)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Murphy, Chris", Range:=Selection.Range)
    oAutoText.Value = "Senator Chris Murphy (D-CT)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Murray, Patty", Range:=Selection.Range)
    oAutoText.Value = "Senator Patty Murray (D-WA)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Nelson, Bill", Range:=Selection.Range)
    oAutoText.Value = "Senator Bill Nelson (D-FL)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Paul, Rand", Range:=Selection.Range)
    oAutoText.Value = "Senator Rand Paul (R-KY)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Portman, Rob", Range:=Selection.Range)
    oAutoText.Value = "Senator Rob Portman (R-OH)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Pryor, Mark", Range:=Selection.Range)
    oAutoText.Value = "Senator Mark Pryor (D-AR)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Reed, Jack", Range:=Selection.Range)
    oAutoText.Value = "Senator Jack Reed (D-RI)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Reid, Harry", Range:=Selection.Range)
    oAutoText.Value = "Senator Harry Reid (D-NV)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Risch, Jim", Range:=Selection.Range)
    oAutoText.Value = "Senator Jim Risch (R-ID)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Roberts, Pat", Range:=Selection.Range)
    oAutoText.Value = "Senator Pat Roberts (R-KS)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Rockefeller, Jay", Range:=Selection.Range)
    oAutoText.Value = "Senator Jay Rockefeller (D-WV)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Rubio, Marco", Range:=Selection.Range)
    oAutoText.Value = "Senator Marco Rubio (R-FL)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Sanders, Bernie", Range:=Selection.Range)
    oAutoText.Value = "Senator Bernie Sanders (I-VT)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Schatz, Brian", Range:=Selection.Range)
    oAutoText.Value = "Senator Brian Schatz (D-HI)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Schumer, Chuck", Range:=Selection.Range)
    oAutoText.Value = "Senator Chuck Schumer (D-NY)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Scott, Tim", Range:=Selection.Range)
    oAutoText.Value = "Senator Tim Scott (R-SC)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Sessions, Jeff", Range:=Selection.Range)
    oAutoText.Value = "Senator Jeff Sessions (R-AL)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Shaheen, Jeanne", Range:=Selection.Range)
    oAutoText.Value = "Senator Jeanne Shaheen (D-NH)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Shelby, Richard", Range:=Selection.Range)
    oAutoText.Value = "Senator Richard Shelby (R-AL)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Stabenow, Debbie", Range:=Selection.Range)
    oAutoText.Value = "Senator Debbie Stabenow (D-MI)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Tester, Jon", Range:=Selection.Range)
    oAutoText.Value = "Senator Jon Tester (D-MT)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Thune, John", Range:=Selection.Range)
    oAutoText.Value = "Senator John Thune (R-SD)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Toomey, Pat", Range:=Selection.Range)
    oAutoText.Value = "Senator Pat Toomey (R-PA)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Udall, Mark", Range:=Selection.Range)
    oAutoText.Value = "Senator Mark Udall (D-CO)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Udall, Tom", Range:=Selection.Range)
    oAutoText.Value = "Senator Tom Udall (D-NM)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Vitter, Dave", Range:=Selection.Range)
    oAutoText.Value = "Senator Dave Vitter (R-LA)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Warner, Mark", Range:=Selection.Range)
    oAutoText.Value = "Senator Mark Warner (D-VA)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Warren, Elizabeth", Range:=Selection.Range)
    oAutoText.Value = "Senator Elizabeth Warren (D-MA)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Whitehouse, Sheldon", Range:=Selection.Range)
    oAutoText.Value = "Senator Sheldon Whitehouse (D-RI)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Wicker, Roger", Range:=Selection.Range)
    oAutoText.Value = "Senator Roger Wicker (R-MS)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Wyden, Ron", Range:=Selection.Range)
    oAutoText.Value = "Senator Ron Wyden (D-OR)"

Set oAutoText = Nothing
End Sub

Private Sub AutoRep1()
Dim oAutoText As AutoTextEntry

Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Abraham Ralph", Range:=Selection.Range)
    oAutoText.Value = "Representative Ralph Abraham (R-LA-05)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Adams Alma", Range:=Selection.Range)
    oAutoText.Value = "Representative Alma Adams (D-NC-12)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Aderholt Robert", Range:=Selection.Range)
    oAutoText.Value = "Representative Robert Aderholt (R-AL-04)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Aguilar Pete", Range:=Selection.Range)
    oAutoText.Value = "Representative Pete Aguilar (D-CA-31)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Allen Rick", Range:=Selection.Range)
    oAutoText.Value = "Representative Rick Allen (R-GA-12)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Amash Justin", Range:=Selection.Range)
    oAutoText.Value = "Representative Justin Amash (R-MI-03)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Amodei Mark", Range:=Selection.Range)
    oAutoText.Value = "Representative Mark Amodei (R-NV-02)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Arrington Jodey", Range:=Selection.Range)
    oAutoText.Value = "Representative Jodey Arrington (R-TX-19)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Babin Brian", Range:=Selection.Range)
    oAutoText.Value = "Representative Brian Babin (R-TX-36)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Bacon Don", Range:=Selection.Range)
    oAutoText.Value = "Representative Don Bacon (R-NE-02)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Banks Jim", Range:=Selection.Range)
    oAutoText.Value = "Representative Jim Banks (R-IN-03)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Barletta Lou", Range:=Selection.Range)
    oAutoText.Value = "Representative Lou Barletta (R-PA-11)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Barr Andy", Range:=Selection.Range)
    oAutoText.Value = "Representative Andy Barr (R-KY-06)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Barragán Nanette", Range:=Selection.Range)
    oAutoText.Value = "Representative Nanette Barragán (D-CA-44)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Barton Joe", Range:=Selection.Range)
    oAutoText.Value = "Representative Joe Barton (R-TX-06)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Bass Karen", Range:=Selection.Range)
    oAutoText.Value = "Representative Karen Bass (D-CA-37)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Beatty Joyce", Range:=Selection.Range)
    oAutoText.Value = "Representative Joyce Beatty (D-OH-03)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Bera Ami", Range:=Selection.Range)
    oAutoText.Value = "Representative Ami Bera (D-CA-07)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Bergman Jack", Range:=Selection.Range)
    oAutoText.Value = "Representative Jack Bergman (R-MI-01)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Beyer Donald Jr.", Range:=Selection.Range)
    oAutoText.Value = "Representative Donald Beyer Jr. (D-VA-08)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Biggs Andy", Range:=Selection.Range)
    oAutoText.Value = "Representative Andy Biggs (R-AZ-05)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Bilirakis Gus", Range:=Selection.Range)
    oAutoText.Value = "Representative Gus Bilirakis (R-FL-12)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Bishop Mike", Range:=Selection.Range)
    oAutoText.Value = "Representative Mike Bishop (R-MI-08)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Bishop Rob", Range:=Selection.Range)
    oAutoText.Value = "Representative Rob Bishop (R-UT-01)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Bishop Sanford Jr.", Range:=Selection.Range)
    oAutoText.Value = "Representative Sanford Bishop Jr. (D-GA-02)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Black Diane", Range:=Selection.Range)
    oAutoText.Value = "Representative Diane Black (R-TN-06)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Blackburn Marsha", Range:=Selection.Range)
    oAutoText.Value = "Representative Marsha Blackburn (R-TN-07)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Blum Rod", Range:=Selection.Range)
    oAutoText.Value = "Representative Rod Blum (R-IA-01)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Blumenauer Earl", Range:=Selection.Range)
    oAutoText.Value = "Representative Earl Blumenauer (D-OR-03)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Blunt Rochester Lisa", Range:=Selection.Range)
    oAutoText.Value = "Representative Lisa Blunt Rochester (D-DE-00)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Bonamici Suzanne", Range:=Selection.Range)
    oAutoText.Value = "Representative Suzanne Bonamici (D-OR-01)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Bordallo Madeleine", Range:=Selection.Range)
    oAutoText.Value = "Representative Madeleine Bordallo (D-GU-00)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Bost Mike", Range:=Selection.Range)
    oAutoText.Value = "Representative Mike Bost (R-IL-12)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Boyle Brendan", Range:=Selection.Range)
    oAutoText.Value = "Representative Brendan Boyle (D-PA-13)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Brady Kevin", Range:=Selection.Range)
    oAutoText.Value = "Representative Kevin Brady (R-TX-08)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Brady Robert", Range:=Selection.Range)
    oAutoText.Value = "Representative Robert Brady (D-PA-01)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Brat Dave", Range:=Selection.Range)
    oAutoText.Value = "Representative Dave Brat (R-VA-07)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Bridenstine Jim", Range:=Selection.Range)
    oAutoText.Value = "Representative Jim Bridenstine (R-OK-01)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Brooks Mo", Range:=Selection.Range)
    oAutoText.Value = "Representative Mo Brooks (R-AL-05)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Brooks Susan", Range:=Selection.Range)
    oAutoText.Value = "Representative Susan Brooks (R-IN-05)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Brown Anthony", Range:=Selection.Range)
    oAutoText.Value = "Representative Anthony Brown (D-MD-04)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Brownley Julia", Range:=Selection.Range)
    oAutoText.Value = "Representative Julia Brownley (D-CA-26)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Buchanan Vern", Range:=Selection.Range)
    oAutoText.Value = "Representative Vern Buchanan (R-FL-16)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Buck Ken", Range:=Selection.Range)
    oAutoText.Value = "Representative Ken Buck (R-CO-04)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Bucshon Larry", Range:=Selection.Range)
    oAutoText.Value = "Representative Larry Bucshon (R-IN-08)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Budd Ted", Range:=Selection.Range)
    oAutoText.Value = "Representative Ted Budd (R-NC-13)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Burgess Michael", Range:=Selection.Range)
    oAutoText.Value = "Representative Michael Burgess (R-TX-26)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Bustos Cheri", Range:=Selection.Range)
    oAutoText.Value = "Representative Cheri Bustos (D-IL-17)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Butterfield G.", Range:=Selection.Range)
    oAutoText.Value = "Representative G. Butterfield (D-NC-01)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Byrne Bradley", Range:=Selection.Range)
    oAutoText.Value = "Representative Bradley Byrne (R-AL-01)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Calvert Ken", Range:=Selection.Range)
    oAutoText.Value = "Representative Ken Calvert (R-CA-42)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Capuano Michael", Range:=Selection.Range)
    oAutoText.Value = "Representative Michael Capuano (D-MA-07)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Carbajal Salud", Range:=Selection.Range)
    oAutoText.Value = "Representative Salud Carbajal (D-CA-24)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Cárdenas Tony", Range:=Selection.Range)
    oAutoText.Value = "Representative Tony Cárdenas (D-CA-29)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Carson André", Range:=Selection.Range)
    oAutoText.Value = "Representative André Carson (D-IN-07)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Carter Earl", Range:=Selection.Range)
    oAutoText.Value = "Representative Earl Carter (R-GA-01)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Carter John", Range:=Selection.Range)
    oAutoText.Value = "Representative John Carter (R-TX-31)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Cartwright Matt", Range:=Selection.Range)
    oAutoText.Value = "Representative Matt Cartwright (D-PA-17)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Castor Kathy", Range:=Selection.Range)
    oAutoText.Value = "Representative Kathy Castor (D-FL-14)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Castro Joaquin", Range:=Selection.Range)
    oAutoText.Value = "Representative Joaquin Castro (D-TX-20)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Chabot Steve", Range:=Selection.Range)
    oAutoText.Value = "Representative Steve Chabot (R-OH-01)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Cheney Liz", Range:=Selection.Range)
    oAutoText.Value = "Representative Liz Cheney (R-WY-00)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Chu Judy", Range:=Selection.Range)
    oAutoText.Value = "Representative Judy Chu (D-CA-27)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Cicilline David", Range:=Selection.Range)
    oAutoText.Value = "Representative David Cicilline (D-RI-01)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Clark Katherine", Range:=Selection.Range)
    oAutoText.Value = "Representative Katherine Clark (D-MA-05)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Clarke Yvette", Range:=Selection.Range)
    oAutoText.Value = "Representative Yvette Clarke (D-NY-09)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Clay Wm.", Range:=Selection.Range)
    oAutoText.Value = "Representative Wm. Clay (D-MO-01)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Cleaver Emanuel", Range:=Selection.Range)
    oAutoText.Value = "Representative Emanuel Cleaver (D-MO-05)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Clyburn James", Range:=Selection.Range)
    oAutoText.Value = "Representative James Clyburn (D-SC-06)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Coffman Mike", Range:=Selection.Range)
    oAutoText.Value = "Representative Mike Coffman (R-CO-06)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Cohen Steve", Range:=Selection.Range)
    oAutoText.Value = "Representative Steve Cohen (D-TN-09)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Cole Tom", Range:=Selection.Range)
    oAutoText.Value = "Representative Tom Cole (R-OK-04)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Collins Chris", Range:=Selection.Range)
    oAutoText.Value = "Representative Chris Collins (R-NY-27)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Collins Doug", Range:=Selection.Range)
    oAutoText.Value = "Representative Doug Collins (R-GA-09)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Comer James", Range:=Selection.Range)
    oAutoText.Value = "Representative James Comer (R-KY-01)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Comstock Barbara", Range:=Selection.Range)
    oAutoText.Value = "Representative Barbara Comstock (R-VA-10)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Conaway K.", Range:=Selection.Range)
    oAutoText.Value = "Representative K. Conaway (R-TX-11)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Connolly Gerald", Range:=Selection.Range)
    oAutoText.Value = "Representative Gerald Connolly (D-VA-11)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Cook Paul", Range:=Selection.Range)
    oAutoText.Value = "Representative Paul Cook (R-CA-08)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Cooper Jim", Range:=Selection.Range)
    oAutoText.Value = "Representative Jim Cooper (D-TN-05)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Correa J.", Range:=Selection.Range)
    oAutoText.Value = "Representative J. Correa (D-CA-46)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Costa Jim", Range:=Selection.Range)
    oAutoText.Value = "Representative Jim Costa (D-CA-16)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Costello Ryan", Range:=Selection.Range)
    oAutoText.Value = "Representative Ryan Costello (R-PA-06)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Courtney Joe", Range:=Selection.Range)
    oAutoText.Value = "Representative Joe Courtney (D-CT-02)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Cramer Kevin", Range:=Selection.Range)
    oAutoText.Value = "Representative Kevin Cramer (R-ND-00)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Crawford Eric", Range:=Selection.Range)
    oAutoText.Value = "Representative Eric Crawford (R-AR-01)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Crist Charlie", Range:=Selection.Range)
    oAutoText.Value = "Representative Charlie Crist (D-FL-13)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Crowley Joseph", Range:=Selection.Range)
    oAutoText.Value = "Representative Joseph Crowley (D-NY-14)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Cuellar Henry", Range:=Selection.Range)
    oAutoText.Value = "Representative Henry Cuellar (D-TX-28)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Culberson John", Range:=Selection.Range)
    oAutoText.Value = "Representative John Culberson (R-TX-07)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Cummings Elijah", Range:=Selection.Range)
    oAutoText.Value = "Representative Elijah Cummings (D-MD-07)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Curbelo Carlos", Range:=Selection.Range)
    oAutoText.Value = "Representative Carlos Curbelo (R-FL-26)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Curtis John", Range:=Selection.Range)
    oAutoText.Value = "Representative John Curtis (R-UT-03)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Davidson Warren", Range:=Selection.Range)
    oAutoText.Value = "Representative Warren Davidson (R-OH-08)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Davis Danny", Range:=Selection.Range)
    oAutoText.Value = "Representative Danny Davis (D-IL-07)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Davis Rodney", Range:=Selection.Range)
    oAutoText.Value = "Representative Rodney Davis (R-IL-13)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Davis Susan", Range:=Selection.Range)
    oAutoText.Value = "Representative Susan Davis (D-CA-53)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="DeFazio Peter", Range:=Selection.Range)
    oAutoText.Value = "Representative Peter DeFazio (D-OR-04)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="DeGette Diana", Range:=Selection.Range)
    oAutoText.Value = "Representative Diana DeGette (D-CO-01)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Delaney John", Range:=Selection.Range)
    oAutoText.Value = "Representative John Delaney (D-MD-06)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="DeLauro Rosa", Range:=Selection.Range)
    oAutoText.Value = "Representative Rosa DeLauro (D-CT-03)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="DelBene Suzan", Range:=Selection.Range)
    oAutoText.Value = "Representative Suzan DelBene (D-WA-01)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Demings Val", Range:=Selection.Range)
    oAutoText.Value = "Representative Val Demings (D-FL-10)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Denham Jeff", Range:=Selection.Range)
    oAutoText.Value = "Representative Jeff Denham (R-CA-10)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Dent Charles", Range:=Selection.Range)
    oAutoText.Value = "Representative Charles Dent (R-PA-15)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="DeSantis Ron", Range:=Selection.Range)
    oAutoText.Value = "Representative Ron DeSantis (R-FL-06)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="DeSaulnier Mark", Range:=Selection.Range)
    oAutoText.Value = "Representative Mark DeSaulnier (D-CA-11)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="DesJarlais Scott", Range:=Selection.Range)
    oAutoText.Value = "Representative Scott DesJarlais (R-TN-04)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Deutch Theodore", Range:=Selection.Range)
    oAutoText.Value = "Representative Theodore Deutch (D-FL-22)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Diaz-Balart Mario", Range:=Selection.Range)
    oAutoText.Value = "Representative Mario Diaz-Balart (R-FL-25)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Dingell Debbie", Range:=Selection.Range)
    oAutoText.Value = "Representative Debbie Dingell (D-MI-12)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Doggett Lloyd", Range:=Selection.Range)
    oAutoText.Value = "Representative Lloyd Doggett (D-TX-35)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Donovan Daniel Jr.", Range:=Selection.Range)
    oAutoText.Value = "Representative Daniel Donovan Jr. (R-NY-11)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Doyle Michael", Range:=Selection.Range)
    oAutoText.Value = "Representative Michael Doyle (D-PA-14)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Duffy Sean", Range:=Selection.Range)
    oAutoText.Value = "Representative Sean Duffy (R-WI-07)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Duncan Jeff", Range:=Selection.Range)
    oAutoText.Value = "Representative Jeff Duncan (R-SC-03)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Duncan John Jr.", Range:=Selection.Range)
    oAutoText.Value = "Representative John Duncan Jr. (R-TN-02)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Dunn Neal", Range:=Selection.Range)
    oAutoText.Value = "Representative Neal Dunn (R-FL-02)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Ellison Keith", Range:=Selection.Range)
    oAutoText.Value = "Representative Keith Ellison (D-MN-05)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Emmer Tom", Range:=Selection.Range)
    oAutoText.Value = "Representative Tom Emmer (R-MN-06)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Engel Eliot", Range:=Selection.Range)
    oAutoText.Value = "Representative Eliot Engel (D-NY-16)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Eshoo Anna", Range:=Selection.Range)
    oAutoText.Value = "Representative Anna Eshoo (D-CA-18)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Espaillat Adriano", Range:=Selection.Range)
    oAutoText.Value = "Representative Adriano Espaillat (D-NY-13)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Estes Ron", Range:=Selection.Range)
    oAutoText.Value = "Representative Ron Estes (R-KS-04)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Esty Elizabeth", Range:=Selection.Range)
    oAutoText.Value = "Representative Elizabeth Esty (D-CT-05)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Evans Dwight", Range:=Selection.Range)
    oAutoText.Value = "Representative Dwight Evans (D-PA-02)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Farenthold Blake", Range:=Selection.Range)
    oAutoText.Value = "Representative Blake Farenthold (R-TX-27)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Faso John", Range:=Selection.Range)
    oAutoText.Value = "Representative John Faso (R-NY-19)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Ferguson Drew IV", Range:=Selection.Range)
    oAutoText.Value = "Representative Drew Ferguson IV (R-GA-03)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Fitzpatrick Brian", Range:=Selection.Range)
    oAutoText.Value = "Representative Brian Fitzpatrick (R-PA-08)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Fleischmann Charles", Range:=Selection.Range)
    oAutoText.Value = "Representative Charles Fleischmann (R-TN-03)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Flores Bill", Range:=Selection.Range)
    oAutoText.Value = "Representative Bill Flores (R-TX-17)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Fortenberry Jeff", Range:=Selection.Range)
    oAutoText.Value = "Representative Jeff Fortenberry (R-NE-01)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Foster Bill", Range:=Selection.Range)
    oAutoText.Value = "Representative Bill Foster (D-IL-11)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Foxx Virginia", Range:=Selection.Range)
    oAutoText.Value = "Representative Virginia Foxx (R-NC-05)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Frankel Lois", Range:=Selection.Range)
    oAutoText.Value = "Representative Lois Frankel (D-FL-21)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Frelinghuysen Rodney", Range:=Selection.Range)
    oAutoText.Value = "Representative Rodney Frelinghuysen (R-NJ-11)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Fudge Marcia", Range:=Selection.Range)
    oAutoText.Value = "Representative Marcia Fudge (D-OH-11)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Gabbard Tulsi", Range:=Selection.Range)
    oAutoText.Value = "Representative Tulsi Gabbard (D-HI-02)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Gaetz Matt", Range:=Selection.Range)
    oAutoText.Value = "Representative Matt Gaetz (R-FL-01)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Gallagher Mike", Range:=Selection.Range)
    oAutoText.Value = "Representative Mike Gallagher (R-WI-08)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Gallego Ruben", Range:=Selection.Range)
    oAutoText.Value = "Representative Ruben Gallego (D-AZ-07)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Garamendi John", Range:=Selection.Range)
    oAutoText.Value = "Representative John Garamendi (D-CA-03)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Garrett Thomas Jr.", Range:=Selection.Range)
    oAutoText.Value = "Representative Thomas Garrett Jr. (R-VA-05)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Gianforte Greg", Range:=Selection.Range)
    oAutoText.Value = "Representative Greg Gianforte (R-MT-00)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Gibbs Bob", Range:=Selection.Range)
    oAutoText.Value = "Representative Bob Gibbs (R-OH-07)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Gohmert Louie", Range:=Selection.Range)
    oAutoText.Value = "Representative Louie Gohmert (R-TX-01)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Gomez Jimmy", Range:=Selection.Range)
    oAutoText.Value = "Representative Jimmy Gomez (D-CA-34)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Gonzalez Vicente", Range:=Selection.Range)
    oAutoText.Value = "Representative Vicente Gonzalez (D-TX-15)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="González-Colón Jenniffer", Range:=Selection.Range)
    oAutoText.Value = "Representative Jenniffer González-Colón (R-PR-00)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Goodlatte Bob", Range:=Selection.Range)
    oAutoText.Value = "Representative Bob Goodlatte (R-VA-06)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Gosar Paul", Range:=Selection.Range)
    oAutoText.Value = "Representative Paul Gosar (R-AZ-04)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Gottheimer Josh", Range:=Selection.Range)
    oAutoText.Value = "Representative Josh Gottheimer (D-NJ-05)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Gowdy Trey", Range:=Selection.Range)
    oAutoText.Value = "Representative Trey Gowdy (R-SC-04)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Granger Kay", Range:=Selection.Range)
    oAutoText.Value = "Representative Kay Granger (R-TX-12)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Graves Garret", Range:=Selection.Range)
    oAutoText.Value = "Representative Garret Graves (R-LA-06)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Graves Sam", Range:=Selection.Range)
    oAutoText.Value = "Representative Sam Graves (R-MO-06)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Graves Tom", Range:=Selection.Range)
    oAutoText.Value = "Representative Tom Graves (R-GA-14)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Green Al", Range:=Selection.Range)
    oAutoText.Value = "Representative Al Green (D-TX-09)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Green Gene", Range:=Selection.Range)
    oAutoText.Value = "Representative Gene Green (D-TX-29)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Griffith Morgan", Range:=Selection.Range)
    oAutoText.Value = "Representative Morgan Griffith (R-VA-09)"

Set oAutoText = Nothing
End Sub

Private Sub AutoRep2()
Dim oAutoText As AutoTextEntry

Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Grijalva Raúl", Range:=Selection.Range)
    oAutoText.Value = "Representative Raúl Grijalva (D-AZ-03)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Grothman Glenn", Range:=Selection.Range)
    oAutoText.Value = "Representative Glenn Grothman (R-WI-06)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Guthrie Brett", Range:=Selection.Range)
    oAutoText.Value = "Representative Brett Guthrie (R-KY-02)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Gutiérrez Luis", Range:=Selection.Range)
    oAutoText.Value = "Representative Luis Gutiérrez (D-IL-04)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Hanabusa Colleen", Range:=Selection.Range)
    oAutoText.Value = "Representative Colleen Hanabusa (D-HI-01)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Handel Karen", Range:=Selection.Range)
    oAutoText.Value = "Representative Karen Handel (R-GA-06)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Harper Gregg", Range:=Selection.Range)
    oAutoText.Value = "Representative Gregg Harper (R-MS-03)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Harris Andy", Range:=Selection.Range)
    oAutoText.Value = "Representative Andy Harris (R-MD-01)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Hartzler Vicky", Range:=Selection.Range)
    oAutoText.Value = "Representative Vicky Hartzler (R-MO-04)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Hastings Alcee", Range:=Selection.Range)
    oAutoText.Value = "Representative Alcee Hastings (D-FL-20)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Heck Denny", Range:=Selection.Range)
    oAutoText.Value = "Representative Denny Heck (D-WA-10)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Hensarling Jeb", Range:=Selection.Range)
    oAutoText.Value = "Representative Jeb Hensarling (R-TX-05)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Herrera Beutler Jaime", Range:=Selection.Range)
    oAutoText.Value = "Representative Jaime Herrera Beutler (R-WA-03)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Hice Jody", Range:=Selection.Range)
    oAutoText.Value = "Representative Jody Hice (R-GA-10)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Higgins Brian", Range:=Selection.Range)
    oAutoText.Value = "Representative Brian Higgins (D-NY-26)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Higgins Clay", Range:=Selection.Range)
    oAutoText.Value = "Representative Clay Higgins (R-LA-03)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Hill French", Range:=Selection.Range)
    oAutoText.Value = "Representative French Hill (R-AR-02)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Himes James", Range:=Selection.Range)
    oAutoText.Value = "Representative James Himes (D-CT-04)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Holding George", Range:=Selection.Range)
    oAutoText.Value = "Representative George Holding (R-NC-02)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Hollingsworth Trey", Range:=Selection.Range)
    oAutoText.Value = "Representative Trey Hollingsworth (R-IN-09)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Hoyer Steny", Range:=Selection.Range)
    oAutoText.Value = "Representative Steny Hoyer (D-MD-05)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Hudson Richard", Range:=Selection.Range)
    oAutoText.Value = "Representative Richard Hudson (R-NC-08)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Huffman Jared", Range:=Selection.Range)
    oAutoText.Value = "Representative Jared Huffman (D-CA-02)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Huizenga Bill", Range:=Selection.Range)
    oAutoText.Value = "Representative Bill Huizenga (R-MI-02)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Hultgren Randy", Range:=Selection.Range)
    oAutoText.Value = "Representative Randy Hultgren (R-IL-14)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Hunter Duncan", Range:=Selection.Range)
    oAutoText.Value = "Representative Duncan Hunter (R-CA-50)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Hurd Will", Range:=Selection.Range)
    oAutoText.Value = "Representative Will Hurd (R-TX-23)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Issa Darrell", Range:=Selection.Range)
    oAutoText.Value = "Representative Darrell Issa (R-CA-49)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Jackson Lee Sheila", Range:=Selection.Range)
    oAutoText.Value = "Representative Sheila Jackson Lee (D-TX-18)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Jayapal Pramila", Range:=Selection.Range)
    oAutoText.Value = "Representative Pramila Jayapal (D-WA-07)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Jeffries Hakeem", Range:=Selection.Range)
    oAutoText.Value = "Representative Hakeem Jeffries (D-NY-08)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Jenkins Evan", Range:=Selection.Range)
    oAutoText.Value = "Representative Evan Jenkins (R-WV-03)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Jenkins Lynn", Range:=Selection.Range)
    oAutoText.Value = "Representative Lynn Jenkins (R-KS-02)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Johnson Bill", Range:=Selection.Range)
    oAutoText.Value = "Representative Bill Johnson (R-OH-06)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Johnson Eddie", Range:=Selection.Range)
    oAutoText.Value = "Representative Eddie Johnson (D-TX-30)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Johnson Henry Jr.", Range:=Selection.Range)
    oAutoText.Value = "Representative Henry Johnson Jr. (D-GA-04)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Johnson Mike", Range:=Selection.Range)
    oAutoText.Value = "Representative Mike Johnson (R-LA-04)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Johnson Sam", Range:=Selection.Range)
    oAutoText.Value = "Representative Sam Johnson (R-TX-03)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Jones Walter", Range:=Selection.Range)
    oAutoText.Value = "Representative Walter Jones (R-NC-03)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Jordan Jim", Range:=Selection.Range)
    oAutoText.Value = "Representative Jim Jordan (R-OH-04)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Joyce David", Range:=Selection.Range)
    oAutoText.Value = "Representative David Joyce (R-OH-14)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Kaptur Marcy", Range:=Selection.Range)
    oAutoText.Value = "Representative Marcy Kaptur (D-OH-09)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Katko John", Range:=Selection.Range)
    oAutoText.Value = "Representative John Katko (R-NY-24)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Keating William", Range:=Selection.Range)
    oAutoText.Value = "Representative William Keating (D-MA-09)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Kelly Mike", Range:=Selection.Range)
    oAutoText.Value = "Representative Mike Kelly (R-PA-03)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Kelly Robin", Range:=Selection.Range)
    oAutoText.Value = "Representative Robin Kelly (D-IL-02)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Kelly Trent", Range:=Selection.Range)
    oAutoText.Value = "Representative Trent Kelly (R-MS-01)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Kennedy Joseph III", Range:=Selection.Range)
    oAutoText.Value = "Representative Joseph Kennedy III (D-MA-04)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Khanna Ro", Range:=Selection.Range)
    oAutoText.Value = "Representative Ro Khanna (D-CA-17)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Kihuen Ruben", Range:=Selection.Range)
    oAutoText.Value = "Representative Ruben Kihuen (D-NV-04)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Kildee Daniel", Range:=Selection.Range)
    oAutoText.Value = "Representative Daniel Kildee (D-MI-05)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Kilmer Derek", Range:=Selection.Range)
    oAutoText.Value = "Representative Derek Kilmer (D-WA-06)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Kind Ron", Range:=Selection.Range)
    oAutoText.Value = "Representative Ron Kind (D-WI-03)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="King Peter", Range:=Selection.Range)
    oAutoText.Value = "Representative Peter King (R-NY-02)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="King Steve", Range:=Selection.Range)
    oAutoText.Value = "Representative Steve King (R-IA-04)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Kinzinger Adam", Range:=Selection.Range)
    oAutoText.Value = "Representative Adam Kinzinger (R-IL-16)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Knight Stephen", Range:=Selection.Range)
    oAutoText.Value = "Representative Stephen Knight (R-CA-25)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Krishnamoorthi Raja", Range:=Selection.Range)
    oAutoText.Value = "Representative Raja Krishnamoorthi (D-IL-08)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Kuster Ann", Range:=Selection.Range)
    oAutoText.Value = "Representative Ann Kuster (D-NH-02)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Kustoff David", Range:=Selection.Range)
    oAutoText.Value = "Representative David Kustoff (R-TN-08)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Labrador Raúl", Range:=Selection.Range)
    oAutoText.Value = "Representative Raúl Labrador (R-ID-01)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="LaHood Darin", Range:=Selection.Range)
    oAutoText.Value = "Representative Darin LaHood (R-IL-18)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="LaMalfa Doug", Range:=Selection.Range)
    oAutoText.Value = "Representative Doug LaMalfa (R-CA-01)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Lamborn Doug", Range:=Selection.Range)
    oAutoText.Value = "Representative Doug Lamborn (R-CO-05)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Lance Leonard", Range:=Selection.Range)
    oAutoText.Value = "Representative Leonard Lance (R-NJ-07)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Langevin James", Range:=Selection.Range)
    oAutoText.Value = "Representative James Langevin (D-RI-02)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Larsen Rick", Range:=Selection.Range)
    oAutoText.Value = "Representative Rick Larsen (D-WA-02)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Larson John", Range:=Selection.Range)
    oAutoText.Value = "Representative John Larson (D-CT-01)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Latta Robert", Range:=Selection.Range)
    oAutoText.Value = "Representative Robert Latta (R-OH-05)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Lawrence Brenda", Range:=Selection.Range)
    oAutoText.Value = "Representative Brenda Lawrence (D-MI-14)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Lawson Al Jr.", Range:=Selection.Range)
    oAutoText.Value = "Representative Al Lawson Jr. (D-FL-05)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Lee Barbara", Range:=Selection.Range)
    oAutoText.Value = "Representative Barbara Lee (D-CA-13)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Levin Sander", Range:=Selection.Range)
    oAutoText.Value = "Representative Sander Levin (D-MI-09)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Lewis Jason", Range:=Selection.Range)
    oAutoText.Value = "Representative Jason Lewis (R-MN-02)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Lewis John", Range:=Selection.Range)
    oAutoText.Value = "Representative John Lewis (D-GA-05)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Lieu Ted", Range:=Selection.Range)
    oAutoText.Value = "Representative Ted Lieu (D-CA-33)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Lipinski Daniel", Range:=Selection.Range)
    oAutoText.Value = "Representative Daniel Lipinski (D-IL-03)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="LoBiondo Frank", Range:=Selection.Range)
    oAutoText.Value = "Representative Frank LoBiondo (R-NJ-02)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Loebsack David", Range:=Selection.Range)
    oAutoText.Value = "Representative David Loebsack (D-IA-02)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Lofgren Zoe", Range:=Selection.Range)
    oAutoText.Value = "Representative Zoe Lofgren (D-CA-19)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Long Billy", Range:=Selection.Range)
    oAutoText.Value = "Representative Billy Long (R-MO-07)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Loudermilk Barry", Range:=Selection.Range)
    oAutoText.Value = "Representative Barry Loudermilk (R-GA-11)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Love Mia", Range:=Selection.Range)
    oAutoText.Value = "Representative Mia Love (R-UT-04)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Lowenthal Alan", Range:=Selection.Range)
    oAutoText.Value = "Representative Alan Lowenthal (D-CA-47)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Lowey Nita", Range:=Selection.Range)
    oAutoText.Value = "Representative Nita Lowey (D-NY-17)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Lucas Frank", Range:=Selection.Range)
    oAutoText.Value = "Representative Frank Lucas (R-OK-03)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Luetkemeyer Blaine", Range:=Selection.Range)
    oAutoText.Value = "Representative Blaine Luetkemeyer (R-MO-03)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Luján Ben", Range:=Selection.Range)
    oAutoText.Value = "Representative Ben Luján (D-NM-03)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Lujan Grisham Michelle", Range:=Selection.Range)
    oAutoText.Value = "Representative Michelle Lujan Grisham (D-NM-01)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Lynch Stephen", Range:=Selection.Range)
    oAutoText.Value = "Representative Stephen Lynch (D-MA-08)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="MacArthur Thomas", Range:=Selection.Range)
    oAutoText.Value = "Representative Thomas MacArthur (R-NJ-03)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Maloney Carolyn", Range:=Selection.Range)
    oAutoText.Value = "Representative Carolyn Maloney (D-NY-12)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Maloney Sean Patrick", Range:=Selection.Range)
    oAutoText.Value = "Representative Sean Patrick Maloney (D-NY-18)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Marchant Kenny", Range:=Selection.Range)
    oAutoText.Value = "Representative Kenny Marchant (R-TX-24)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Marino Tom", Range:=Selection.Range)
    oAutoText.Value = "Representative Tom Marino (R-PA-10)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Marshall Roger", Range:=Selection.Range)
    oAutoText.Value = "Representative Roger Marshall (R-KS-01)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Massie Thomas", Range:=Selection.Range)
    oAutoText.Value = "Representative Thomas Massie (R-KY-04)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Mast Brian", Range:=Selection.Range)
    oAutoText.Value = "Representative Brian Mast (R-FL-18)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Matsui Doris", Range:=Selection.Range)
    oAutoText.Value = "Representative Doris Matsui (D-CA-06)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="McCarthy Kevin", Range:=Selection.Range)
    oAutoText.Value = "Representative Kevin McCarthy (R-CA-23)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="McCaul Michael", Range:=Selection.Range)
    oAutoText.Value = "Representative Michael McCaul (R-TX-10)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="McClintock Tom", Range:=Selection.Range)
    oAutoText.Value = "Representative Tom McClintock (R-CA-04)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="McCollum Betty", Range:=Selection.Range)
    oAutoText.Value = "Representative Betty McCollum (D-MN-04)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="McEachin Donald", Range:=Selection.Range)
    oAutoText.Value = "Representative Donald McEachin (D-VA-04)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="McGovern James", Range:=Selection.Range)
    oAutoText.Value = "Representative James McGovern (D-MA-02)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="McHenry Patrick", Range:=Selection.Range)
    oAutoText.Value = "Representative Patrick McHenry (R-NC-10)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="McKinley David", Range:=Selection.Range)
    oAutoText.Value = "Representative David McKinley (R-WV-01)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="McMorris Rodgers Cathy", Range:=Selection.Range)
    oAutoText.Value = "Representative Cathy McMorris Rodgers (R-WA-05)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="McNerney Jerry", Range:=Selection.Range)
    oAutoText.Value = "Representative Jerry McNerney (D-CA-09)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="McSally Martha", Range:=Selection.Range)
    oAutoText.Value = "Representative Martha McSally (R-AZ-02)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Meadows Mark", Range:=Selection.Range)
    oAutoText.Value = "Representative Mark Meadows (R-NC-11)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Meehan Patrick", Range:=Selection.Range)
    oAutoText.Value = "Representative Patrick Meehan (R-PA-07)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Meeks Gregory", Range:=Selection.Range)
    oAutoText.Value = "Representative Gregory Meeks (D-NY-05)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Meng Grace", Range:=Selection.Range)
    oAutoText.Value = "Representative Grace Meng (D-NY-06)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Messer Luke", Range:=Selection.Range)
    oAutoText.Value = "Representative Luke Messer (R-IN-06)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Mitchell Paul", Range:=Selection.Range)
    oAutoText.Value = "Representative Paul Mitchell (R-MI-10)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Moolenaar John", Range:=Selection.Range)
    oAutoText.Value = "Representative John Moolenaar (R-MI-04)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Mooney Alexander", Range:=Selection.Range)
    oAutoText.Value = "Representative Alexander Mooney (R-WV-02)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Moore Gwen", Range:=Selection.Range)
    oAutoText.Value = "Representative Gwen Moore (D-WI-04)"

Set oAutoText = Nothing
End Sub

Private Sub AutoRep3()
Dim oAutoText As AutoTextEntry

Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Moulton Seth", Range:=Selection.Range)
    oAutoText.Value = "Representative Seth Moulton (D-MA-06)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Mullin Markwayne", Range:=Selection.Range)
    oAutoText.Value = "Representative Markwayne Mullin (R-OK-02)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Murphy Stephanie", Range:=Selection.Range)
    oAutoText.Value = "Representative Stephanie Murphy (D-FL-07)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Nadler Jerrold", Range:=Selection.Range)
    oAutoText.Value = "Representative Jerrold Nadler (D-NY-10)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Napolitano Grace", Range:=Selection.Range)
    oAutoText.Value = "Representative Grace Napolitano (D-CA-32)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Neal Richard", Range:=Selection.Range)
    oAutoText.Value = "Representative Richard Neal (D-MA-01)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Newhouse Dan", Range:=Selection.Range)
    oAutoText.Value = "Representative Dan Newhouse (R-WA-04)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Noem Kristi", Range:=Selection.Range)
    oAutoText.Value = "Representative Kristi Noem (R-SD-00)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Nolan Richard", Range:=Selection.Range)
    oAutoText.Value = "Representative Richard Nolan (D-MN-08)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Norcross Donald", Range:=Selection.Range)
    oAutoText.Value = "Representative Donald Norcross (D-NJ-01)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Norman Ralph", Range:=Selection.Range)
    oAutoText.Value = "Representative Ralph Norman (R-SC-05)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Norton Eleanor", Range:=Selection.Range)
    oAutoText.Value = "Representative Eleanor Norton (D-DC-00)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Nunes Devin", Range:=Selection.Range)
    oAutoText.Value = "Representative Devin Nunes (R-CA-22)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="O'Halleran Tom", Range:=Selection.Range)
    oAutoText.Value = "Representative Tom O'Halleran (D-AZ-01)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Olson Pete", Range:=Selection.Range)
    oAutoText.Value = "Representative Pete Olson (R-TX-22)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="O'Rourke Beto", Range:=Selection.Range)
    oAutoText.Value = "Representative Beto O'Rourke (D-TX-16)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Palazzo Steven", Range:=Selection.Range)
    oAutoText.Value = "Representative Steven Palazzo (R-MS-04)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Pallone Frank Jr.", Range:=Selection.Range)
    oAutoText.Value = "Representative Frank Pallone Jr. (D-NJ-06)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Palmer Gary", Range:=Selection.Range)
    oAutoText.Value = "Representative Gary Palmer (R-AL-06)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Panetta Jimmy", Range:=Selection.Range)
    oAutoText.Value = "Representative Jimmy Panetta (D-CA-20)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Pascrell Bill Jr.", Range:=Selection.Range)
    oAutoText.Value = "Representative Bill Pascrell Jr. (D-NJ-09)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Paulsen Erik", Range:=Selection.Range)
    oAutoText.Value = "Representative Erik Paulsen (R-MN-03)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Payne Donald Jr.", Range:=Selection.Range)
    oAutoText.Value = "Representative Donald Payne Jr. (D-NJ-10)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Pearce Stevan", Range:=Selection.Range)
    oAutoText.Value = "Representative Stevan Pearce (R-NM-02)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Pelosi Nancy", Range:=Selection.Range)
    oAutoText.Value = "Representative Nancy Pelosi (D-CA-12)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Perlmutter Ed", Range:=Selection.Range)
    oAutoText.Value = "Representative Ed Perlmutter (D-CO-07)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Perry Scott", Range:=Selection.Range)
    oAutoText.Value = "Representative Scott Perry (R-PA-04)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Peters Scott", Range:=Selection.Range)
    oAutoText.Value = "Representative Scott Peters (D-CA-52)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Peterson Collin", Range:=Selection.Range)
    oAutoText.Value = "Representative Collin Peterson (D-MN-07)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Pingree Chellie", Range:=Selection.Range)
    oAutoText.Value = "Representative Chellie Pingree (D-ME-01)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Pittenger Robert", Range:=Selection.Range)
    oAutoText.Value = "Representative Robert Pittenger (R-NC-09)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Plaskett Stacey", Range:=Selection.Range)
    oAutoText.Value = "Representative Stacey Plaskett (D-VI-00)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Pocan Mark", Range:=Selection.Range)
    oAutoText.Value = "Representative Mark Pocan (D-WI-02)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Poe Ted", Range:=Selection.Range)
    oAutoText.Value = "Representative Ted Poe (R-TX-02)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Poliquin Bruce", Range:=Selection.Range)
    oAutoText.Value = "Representative Bruce Poliquin (R-ME-02)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Polis Jared", Range:=Selection.Range)
    oAutoText.Value = "Representative Jared Polis (D-CO-02)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Posey Bill", Range:=Selection.Range)
    oAutoText.Value = "Representative Bill Posey (R-FL-08)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Price David", Range:=Selection.Range)
    oAutoText.Value = "Representative David Price (D-NC-04)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Quigley Mike", Range:=Selection.Range)
    oAutoText.Value = "Representative Mike Quigley (D-IL-05)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Radewagen Aumua Amata", Range:=Selection.Range)
    oAutoText.Value = "Representative Aumua Amata Radewagen (R-AS-00)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Raskin Jamie", Range:=Selection.Range)
    oAutoText.Value = "Representative Jamie Raskin (D-MD-08)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Ratcliffe John", Range:=Selection.Range)
    oAutoText.Value = "Representative John Ratcliffe (R-TX-04)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Reed Tom", Range:=Selection.Range)
    oAutoText.Value = "Representative Tom Reed (R-NY-23)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Reichert David", Range:=Selection.Range)
    oAutoText.Value = "Representative David Reichert (R-WA-08)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Renacci James", Range:=Selection.Range)
    oAutoText.Value = "Representative James Renacci (R-OH-16)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Rice Kathleen", Range:=Selection.Range)
    oAutoText.Value = "Representative Kathleen Rice (D-NY-04)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Rice Tom", Range:=Selection.Range)
    oAutoText.Value = "Representative Tom Rice (R-SC-07)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Richmond Cedric", Range:=Selection.Range)
    oAutoText.Value = "Representative Cedric Richmond (D-LA-02)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Roby Martha", Range:=Selection.Range)
    oAutoText.Value = "Representative Martha Roby (R-AL-02)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Roe David", Range:=Selection.Range)
    oAutoText.Value = "Representative David Roe (R-TN-01)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Rogers Harold", Range:=Selection.Range)
    oAutoText.Value = "Representative Harold Rogers (R-KY-05)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Rogers Mike", Range:=Selection.Range)
    oAutoText.Value = "Representative Mike Rogers (R-AL-03)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Rohrabacher Dana", Range:=Selection.Range)
    oAutoText.Value = "Representative Dana Rohrabacher (R-CA-48)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Rokita Todd", Range:=Selection.Range)
    oAutoText.Value = "Representative Todd Rokita (R-IN-04)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Rooney Francis", Range:=Selection.Range)
    oAutoText.Value = "Representative Francis Rooney (R-FL-19)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Rooney Thomas", Range:=Selection.Range)
    oAutoText.Value = "Representative Thomas Rooney (R-FL-17)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Rosen Jacky", Range:=Selection.Range)
    oAutoText.Value = "Representative Jacky Rosen (D-NV-03)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Roskam Peter", Range:=Selection.Range)
    oAutoText.Value = "Representative Peter Roskam (R-IL-06)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Ros-Lehtinen Ileana", Range:=Selection.Range)
    oAutoText.Value = "Representative Ileana Ros-Lehtinen (R-FL-27)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Ross Dennis", Range:=Selection.Range)
    oAutoText.Value = "Representative Dennis Ross (R-FL-15)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Rothfus Keith", Range:=Selection.Range)
    oAutoText.Value = "Representative Keith Rothfus (R-PA-12)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Rouzer David", Range:=Selection.Range)
    oAutoText.Value = "Representative David Rouzer (R-NC-07)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Roybal-Allard Lucille", Range:=Selection.Range)
    oAutoText.Value = "Representative Lucille Roybal-Allard (D-CA-40)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Royce Edward", Range:=Selection.Range)
    oAutoText.Value = "Representative Edward Royce (R-CA-39)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Ruiz Raul", Range:=Selection.Range)
    oAutoText.Value = "Representative Raul Ruiz (D-CA-36)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Ruppersberger Dutch", Range:=Selection.Range)
    oAutoText.Value = "Representative Dutch Ruppersberger (D-MD-02)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Rush Bobby", Range:=Selection.Range)
    oAutoText.Value = "Representative Bobby Rush (D-IL-01)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Russell Steve", Range:=Selection.Range)
    oAutoText.Value = "Representative Steve Russell (R-OK-05)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Rutherford John", Range:=Selection.Range)
    oAutoText.Value = "Representative John Rutherford (R-FL-04)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Ryan Paul", Range:=Selection.Range)
    oAutoText.Value = "Representative Paul Ryan (R-WI-01)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Ryan Tim", Range:=Selection.Range)
    oAutoText.Value = "Representative Tim Ryan (D-OH-13)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Sablan Gregorio", Range:=Selection.Range)
    oAutoText.Value = "Representative Gregorio Sablan (D-MP-00)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Sánchez Linda", Range:=Selection.Range)
    oAutoText.Value = "Representative Linda Sánchez (D-CA-38)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Sanford Mark", Range:=Selection.Range)
    oAutoText.Value = "Representative Mark Sanford (R-SC-01)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Sarbanes John", Range:=Selection.Range)
    oAutoText.Value = "Representative John Sarbanes (D-MD-03)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Scalise Steve", Range:=Selection.Range)
    oAutoText.Value = "Representative Steve Scalise (R-LA-01)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Schakowsky Janice", Range:=Selection.Range)
    oAutoText.Value = "Representative Janice Schakowsky (D-IL-09)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Schiff Adam", Range:=Selection.Range)
    oAutoText.Value = "Representative Adam Schiff (D-CA-28)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Schneider Bradley", Range:=Selection.Range)
    oAutoText.Value = "Representative Bradley Schneider (D-IL-10)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Schrader Kurt", Range:=Selection.Range)
    oAutoText.Value = "Representative Kurt Schrader (D-OR-05)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Schweikert David", Range:=Selection.Range)
    oAutoText.Value = "Representative David Schweikert (R-AZ-06)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Scott Austin", Range:=Selection.Range)
    oAutoText.Value = "Representative Austin Scott (R-GA-08)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Scott David", Range:=Selection.Range)
    oAutoText.Value = "Representative David Scott (D-GA-13)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Scott Robert", Range:=Selection.Range)
    oAutoText.Value = "Representative Robert Scott (D-VA-03)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Sensenbrenner James Jr.", Range:=Selection.Range)
    oAutoText.Value = "Representative James Sensenbrenner Jr. (R-WI-05)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Serrano José", Range:=Selection.Range)
    oAutoText.Value = "Representative José Serrano (D-NY-15)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Sessions Pete", Range:=Selection.Range)
    oAutoText.Value = "Representative Pete Sessions (R-TX-32)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Sewell Terri", Range:=Selection.Range)
    oAutoText.Value = "Representative Terri Sewell (D-AL-07)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Shea-Porter Carol", Range:=Selection.Range)
    oAutoText.Value = "Representative Carol Shea-Porter (D-NH-01)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Sherman Brad", Range:=Selection.Range)
    oAutoText.Value = "Representative Brad Sherman (D-CA-30)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Shimkus John", Range:=Selection.Range)
    oAutoText.Value = "Representative John Shimkus (R-IL-15)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Shuster Bill", Range:=Selection.Range)
    oAutoText.Value = "Representative Bill Shuster (R-PA-09)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Simpson Michael", Range:=Selection.Range)
    oAutoText.Value = "Representative Michael Simpson (R-ID-02)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Sinema Kyrsten", Range:=Selection.Range)
    oAutoText.Value = "Representative Kyrsten Sinema (D-AZ-09)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Sires Albio", Range:=Selection.Range)
    oAutoText.Value = "Representative Albio Sires (D-NJ-08)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Slaughter Louise", Range:=Selection.Range)
    oAutoText.Value = "Representative Louise Slaughter (D-NY-25)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Smith Adam", Range:=Selection.Range)
    oAutoText.Value = "Representative Adam Smith (D-WA-09)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Smith Adrian", Range:=Selection.Range)
    oAutoText.Value = "Representative Adrian Smith (R-NE-03)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Smith Christopher", Range:=Selection.Range)
    oAutoText.Value = "Representative Christopher Smith (R-NJ-04)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Smith Jason", Range:=Selection.Range)
    oAutoText.Value = "Representative Jason Smith (R-MO-08)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Smith Lamar", Range:=Selection.Range)
    oAutoText.Value = "Representative Lamar Smith (R-TX-21)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Smucker Lloyd", Range:=Selection.Range)
    oAutoText.Value = "Representative Lloyd Smucker (R-PA-16)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Soto Darren", Range:=Selection.Range)
    oAutoText.Value = "Representative Darren Soto (D-FL-09)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Speier Jackie", Range:=Selection.Range)
    oAutoText.Value = "Representative Jackie Speier (D-CA-14)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Stefanik Elise", Range:=Selection.Range)
    oAutoText.Value = "Representative Elise Stefanik (R-NY-21)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Stewart Chris", Range:=Selection.Range)
    oAutoText.Value = "Representative Chris Stewart (R-UT-02)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Stivers Steve", Range:=Selection.Range)
    oAutoText.Value = "Representative Steve Stivers (R-OH-15)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Suozzi Thomas", Range:=Selection.Range)
    oAutoText.Value = "Representative Thomas Suozzi (D-NY-03)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Swalwell Eric", Range:=Selection.Range)
    oAutoText.Value = "Representative Eric Swalwell (D-CA-15)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Takano Mark", Range:=Selection.Range)
    oAutoText.Value = "Representative Mark Takano (D-CA-41)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Taylor Scott", Range:=Selection.Range)
    oAutoText.Value = "Representative Scott Taylor (R-VA-02)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Tenney Claudia", Range:=Selection.Range)
    oAutoText.Value = "Representative Claudia Tenney (R-NY-22)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Thompson Bennie", Range:=Selection.Range)
    oAutoText.Value = "Representative Bennie Thompson (D-MS-02)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Thompson Glenn", Range:=Selection.Range)
    oAutoText.Value = "Representative Glenn Thompson (R-PA-05)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Thompson Mike", Range:=Selection.Range)
    oAutoText.Value = "Representative Mike Thompson (D-CA-05)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Thornberry Mac", Range:=Selection.Range)
    oAutoText.Value = "Representative Mac Thornberry (R-TX-13)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Tipton Scott", Range:=Selection.Range)
    oAutoText.Value = "Representative Scott Tipton (R-CO-03)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Titus Dina", Range:=Selection.Range)
    oAutoText.Value = "Representative Dina Titus (D-NV-01)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Tonko Paul", Range:=Selection.Range)
    oAutoText.Value = "Representative Paul Tonko (D-NY-20)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Torres Norma", Range:=Selection.Range)
    oAutoText.Value = "Representative Norma Torres (D-CA-35)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Trott David", Range:=Selection.Range)
    oAutoText.Value = "Representative David Trott (R-MI-11)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Tsongas Niki", Range:=Selection.Range)
    oAutoText.Value = "Representative Niki Tsongas (D-MA-03)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Turner Michael", Range:=Selection.Range)
    oAutoText.Value = "Representative Michael Turner (R-OH-10)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Upton Fred", Range:=Selection.Range)
    oAutoText.Value = "Representative Fred Upton (R-MI-06)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Valadao David", Range:=Selection.Range)
    oAutoText.Value = "Representative David Valadao (R-CA-21)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Vargas Juan", Range:=Selection.Range)
    oAutoText.Value = "Representative Juan Vargas (D-CA-51)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Veasey Marc", Range:=Selection.Range)
    oAutoText.Value = "Representative Marc Veasey (D-TX-33)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Vela Filemon", Range:=Selection.Range)
    oAutoText.Value = "Representative Filemon Vela (D-TX-34)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Velázquez Nydia", Range:=Selection.Range)
    oAutoText.Value = "Representative Nydia Velázquez (D-NY-07)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Visclosky Peter", Range:=Selection.Range)
    oAutoText.Value = "Representative Peter Visclosky (D-IN-01)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Wagner Ann", Range:=Selection.Range)
    oAutoText.Value = "Representative Ann Wagner (R-MO-02)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Walberg Tim", Range:=Selection.Range)
    oAutoText.Value = "Representative Tim Walberg (R-MI-07)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Walden Greg", Range:=Selection.Range)
    oAutoText.Value = "Representative Greg Walden (R-OR-02)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Walker Mark", Range:=Selection.Range)
    oAutoText.Value = "Representative Mark Walker (R-NC-06)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Walorski Jackie", Range:=Selection.Range)
    oAutoText.Value = "Representative Jackie Walorski (R-IN-02)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Walters Mimi", Range:=Selection.Range)
    oAutoText.Value = "Representative Mimi Walters (R-CA-45)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Walz Timothy", Range:=Selection.Range)
    oAutoText.Value = "Representative Timothy Walz (D-MN-01)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Wasserman Schultz Debbie", Range:=Selection.Range)
    oAutoText.Value = "Representative Debbie Wasserman Schultz (D-FL-23)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Waters Maxine", Range:=Selection.Range)
    oAutoText.Value = "Representative Maxine Waters (D-CA-43)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Watson Coleman Bonnie", Range:=Selection.Range)
    oAutoText.Value = "Representative Bonnie Watson Coleman (D-NJ-12)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Weber Randy Sr.", Range:=Selection.Range)
    oAutoText.Value = "Representative Randy Weber Sr. (R-TX-14)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Webster Daniel", Range:=Selection.Range)
    oAutoText.Value = "Representative Daniel Webster (R-FL-11)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Welch Peter", Range:=Selection.Range)
    oAutoText.Value = "Representative Peter Welch (D-VT-00)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Wenstrup Brad", Range:=Selection.Range)
    oAutoText.Value = "Representative Brad Wenstrup (R-OH-02)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Westerman Bruce", Range:=Selection.Range)
    oAutoText.Value = "Representative Bruce Westerman (R-AR-04)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Williams Roger", Range:=Selection.Range)
    oAutoText.Value = "Representative Roger Williams (R-TX-25)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Wilson Frederica", Range:=Selection.Range)
    oAutoText.Value = "Representative Frederica Wilson (D-FL-24)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Wilson Joe", Range:=Selection.Range)
    oAutoText.Value = "Representative Joe Wilson (R-SC-02)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Wittman Robert", Range:=Selection.Range)
    oAutoText.Value = "Representative Robert Wittman (R-VA-01)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Womack Steve", Range:=Selection.Range)
    oAutoText.Value = "Representative Steve Womack (R-AR-03)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Woodall Rob", Range:=Selection.Range)
    oAutoText.Value = "Representative Rob Woodall (R-GA-07)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Yarmuth John", Range:=Selection.Range)
    oAutoText.Value = "Representative John Yarmuth (D-KY-03)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Yoder Kevin", Range:=Selection.Range)
    oAutoText.Value = "Representative Kevin Yoder (R-KS-03)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Yoho Ted", Range:=Selection.Range)
    oAutoText.Value = "Representative Ted Yoho (R-FL-03)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Young David", Range:=Selection.Range)
    oAutoText.Value = "Representative David Young (R-IA-03)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Young Don", Range:=Selection.Range)
    oAutoText.Value = "Representative Don Young (R-AK-00)"
Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:="Zeldin Lee", Range:=Selection.Range)
    oAutoText.Value = "Representative Lee Zeldin (R-NY-01)"

Set oAutoText = Nothing
End Sub
