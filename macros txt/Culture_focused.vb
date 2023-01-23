'MacroName:Culture_focused
'MacroDescription:Connexion macro for adding LGBTQ+ identity related subjects.
'January, 2023

Sub Main
   
Dim CS As Object
Dim sData as String
dim sd
sd = chr(223) 
dim nd
nd = chr$(9)   
    
On Error Resume Next
Set CS = GetObject(,"Connex.Client")
On Error GoTo 0
If CS Is Nothing Then
   Set CS = CreateObject("Connex.Client")
End If


dim aa$
aa$ = " "
dim ba$
ba$ = " "
dim ca$
ca$ = " "
dim da$
da$ = " "
dim ea$
ea$ = " "
dim fa$
fa$ = " "
dim ga$
ga$ = " "
dim ja$
ja$ = " "
dim na$
na$ = " "

'****************************************************************************************
' Drop-down menu selections/lists
'****************************************************************************************

dim Subcultures$
Subcultures$=" "+nd+ "Subcultures"+nd+"  Ball culture"+nd+"    Ballroom families"+nd+"  Dandyism"+nd+"  Decadence" _
+nd+"  Fandom"+nd+"    NT > Fandom"+nd+"  Gym subculture"+nd+"  Leather community"+nd+"    NT > Leather community" _
+nd+"  LGBTQ+ culture"+nd+"    Bisexual culture"+nd+"    Gay culture"+nd+"      NT > Gay culture"+nd+"    Lesbian culture" _
+nd+"      NT > Lesbian culture"+nd+"    LGBTQ+ arts"+nd+"      LGBTQ+ cinema"+nd+"        NT > LGBTQ+ cinema"+nd+"    LGBTQ+ literature" _
+nd+"      NT > LGBTQ+ literature"+nd+"      Bisexual literature"+nd+"      Gay literature"+nd+"      Lesbian literature"+nd+"      LGBTQ+ drama" _
+nd+"      LGBTQ+ fiction"+nd+"      LGBTQ+ poetry"+nd+"      LGBTQ+ youth literature"+nd+"      Queer literature"+nd+"      Transgender literature" _
+nd+"    Queer culture"+nd+"    Transgender culture"+nd+"      Transgender slang"+nd+"  Rubber subculture"+nd+"    Rubber sex" _
+nd+"    Rubbermen"+nd+"    Rubberwomen"

dim LGBTQplus_awards$
LGBTQplus_awards$=" "+nd+ "LGBTQ+ awards"+nd+"LGBTQ+ literary awards"

dim Body_adornment$
Body_adornment$=" "+nd+ "Body adornment"+nd+"Piercings"

dim Clothing$
Clothing$=" "+nd+ "Clothing"+nd+"Costumes"+nd+"Lingerie"

dim LGBTQplus_events$
LGBTQplus_events$=" "+nd+ "LGBTQ+ events"+nd+"Drag balls"+nd+"Dyke marches"+nd+"Gay pride"+nd+"    Gay pride week" _
+nd+"Halloween (LGBTQ+ culture)"+nd+"Lesbian pride"+nd+"LGBTQ+ sporting events"+nd+"Trans marches"+nd+"Women's music festivals" _

dim LGBTQplus_slang$
LGBTQplus_slang$=" "+nd+ "LGBTQ+ slang"+nd+"Bisexual slang"+nd+"Gay slang"+nd+"Lesbian slang"+nd+"Poppers" _
+nd+"Queer slang"+nd+"Reclaimed terms"+nd+"  Crip"+nd+"  Dykes"+nd+"  Faggots" _
+nd+"  Queer people"+nd+"Transgender slang"

dim LGBTQplus_symbols$
LGBTQplus_symbols$=" "+nd+ "LGBTQ+ symbols"+nd+"Freedom rings"+nd+"Handkerchief codes"+nd+"Labrys"+nd+"Lambda" _
+nd+"Pink triangles"+nd+"Rainbow flags"+nd+"Red ribbons (AIDS)"

dim LGBTQplus_clubs$
LGBTQplus_clubs$=" "+nd+ "LGBTQ+ clubs"+nd+"LGBTQ+ book clubs"+nd+"LGBTQ+ motorcycle clubs"+nd+"LGBTQ+ sports clubs"+nd+"  Gay sports clubs" _
+nd+"  Lesbian sports clubs"+nd+"LGBTQ+ students' clubs"+nd+"Sex clubs"

dim LGBTQplus_night_life$
LGBTQplus_night_life$=" "+nd+ "LGBTQ+ night life"+nd+"Gay discos"+nd+"Gay saunas"+nd+"  Glory holes"

dim Miscellanious_culture$
Miscellanious_culture$=" "+nd+ "Camp (Gay culture)"+nd+"Gay sensibility"+nd+"Gay theater groups"+nd+"Gaydar"+nd+"Lesbian sensibility" _
+nd+"Lesbian theater groups"+nd+"LGBTQ+ death notices"+nd+"LGBTQ+ film festivals"+nd+"LGBTQ+ monuments"+nd+"LGBTQ+ older people's organisations" _
+nd+"LGBTQ+ weddings"+nd+"Musicals"+nd+"Nudism"+nd+"Sex scandals"+nd+"Social privilege" _
+nd+"Voguing"


'****************************************************************************************
' Menu layout and Titles
'****************************************************************************************

Begin Dialog MFHL 400, 200, "Choose the Homosaurus term you want to use..."

Text 14,7,95,135, "Subcultures"
   DropListBox 200,20,130,135, Subcultures$, .Subculturessub
Text 200,7,80,14, "LGBTQ+ awards"
   DropListBox 200, 20, 130, 135, LGBTQplus_awards$, .LGBTQplus_awardssub

Text 14,45,120,14, "Body adornment"
   DropListBox 14, 58, 130, 200, Body_adornment$, .Body_adornmentsub
Text 200,45,200,14, "Clothing"
   DropListBox 200, 58, 130, 135, Clothing$, .Clothingsub

Text 14,83,95,135, "LGBTQ+ events"
   DropListBox 14, 96, 130, 135, LGBTQplus_events$, .LGBTQplus_eventssub
Text 200,83,80,14, "LGBTQ+ slang"
   DropListBox 200, 96, 130, 135, LGBTQplus_slang$, .LGBTQplus_slangsub

Text 14,121,95,135, "LGBTQ+ symbols"
   DropListBox 14, 135, 130, 135, LGBTQplus_symbols$, .LGBTQplus_symbolssub
Text 200,121,80,14, "LGBTQ+ clubs"
   DropListBox 200, 135, 130, 135, LGBTQplus_clubs$, .LGBTQplus_clubssub

Text 14,159,130,135, "LGBTQ+ night life"
   DropListBox 14,172,130,130, LGBTQplus_night_life$, .LGBTQplus_night_lifesub
Text 200,159,130,135, "Terms that are not in a hierarchy"
   DropListBox 200, 172, 130, 135, Miscellanious_culture$, .Miscellanious_culturesub


Button 64,150,40,20, "I'm done",    .fin
CancelButton 130,150,40,20

End Dialog
dim quick as MFHL
On Error Goto Cancelled
dialog quick
   
'****************************************************************************************
' 650_7 strings for bib records
'****************************************************************************************
If quick.Subculturessub=0 Then na$ = " "
If quick.Subculturessub=1 Then na$ = "Subcultures" & " " & sd & "2 homoit" & " " & sd & "0 https://homosaurus.org/v3/homoit0001345"
If quick.Subculturessub=2 Then na$ = "Ball culture" & " " & sd & "2 homoit" & " " & sd & "0 https://homosaurus.org/v3/homoit0000092"
If quick.Subculturessub=3 Then na$ = "Ballroom families" & " " & sd & "2 homoit" & " " & sd & "0 https://homosaurus.org/v3/homoit0000094"
If quick.Subculturessub=4 Then na$ = "Dandyism" & " " & sd & "2 homoit" & " " & sd & "0 https://homosaurus.org/v3/homoit0000327"
If quick.Subculturessub=5 Then na$ = "Decadence" & " " & sd & "2 homoit" & " " & sd & "0 https://homosaurus.org/v3/homoit0000338"
If quick.Subculturessub=6 Then na$ = "Fandom" & " " & sd & "2 homoit" & " " & sd & "0 https://homosaurus.org/v3/homoit0002156"
If quick.Subculturessub=7 Then CS.RunMacro "Homosaurus_subjects.mbk!Fandom"  
If quick.Subculturessub=8 Then na$ = "Gym subculture" & " " & sd & "2 homoit" & " " & sd & "0 https://homosaurus.org/v3/homoit0001654"
If quick.Subculturessub=9 Then na$ = "Leather community" & " " & sd & "2 homoit" & " " & sd & "0 https://homosaurus.org/v3/homoit0000705"
If quick.Subculturessub=10 Then CS.RunMacro "Homosaurus_subjects.mbk!Leather_community"
If quick.Subculturessub=11 Then na$ = "LGBTQ+ culture" & " " & sd & "2 homoit" & " " & sd & "0 https://homosaurus.org/v3/homoit0001831"
If quick.Subculturessub=12 Then na$ = "Bisexual culture" & " " & sd & "2 homoit" & " " & sd & "0 https://homosaurus.org/v3/homoit0001833"
If quick.Subculturessub=13 Then na$ = "Gay culture" & " " & sd & "2 homoit" & " " & sd & "0 https://homosaurus.org/v3/homoit0000487"
If quick.Subculturessub=14 Then CS.RunMacro "Homosaurus_subjects.mbk!Gay_culture"
If quick.Subculturessub=15 Then na$ = "Lesbian culture" & " " & sd & "2 homoit" & " " & sd & "0 https://homosaurus.org/v3/homoit0000382"
If quick.Subculturessub=16 Then CS.RunMacro "Homosaurus_subjects.mbk!Lesbian_culture"
If quick.Subculturessub=17 Then na$ = "LGBTQ+ arts" & " " & sd & "2 homoit" & " " & sd & "0 https://homosaurus.org/v3/homoit0000816"
If quick.Subculturessub=18 Then na$ = "LGBTQ+ cinema" & " " & sd & "2 homoit" & " " & sd & "0 https://homosaurus.org/v3/homoit0002327"
If quick.Subculturessub=19 Then CS.RunMacro "Homosaurus_subjects.mbk!LGBTQ_cinema"
If quick.Subculturessub=20 Then na$ = "LGBTQ+ literature" & " " & sd & "2 homoit" & " " & sd & "0 https://homosaurus.org/v3/homoit0000972"
If quick.Subculturessub=21 Then CS.RunMacro "Homosaurus_subjects.mbk!LGBTQ_literature"
If quick.Subculturessub=22 Then na$ = "Bisexual literature" & " " & sd & "2 homoit" & " " & sd & "0 https://homosaurus.org/v3/homoit0000157"
If quick.Subculturessub=23 Then na$ = "Gay literature" & " " & sd & "2 homoit" & " " & sd & "0 https://homosaurus.org/v3/homoit0000502"
If quick.Subculturessub=24 Then na$ = "Lesbian literature" & " " & sd & "2 homoit" & " " & sd & "0 https://homosaurus.org/v3/homoit0000750"
If quick.Subculturessub=25 Then na$ = "LGBTQ+ drama" & " " & sd & "2 homoit" & " " & sd & "0 https://homosaurus.org/v3/homoit0000376"
If quick.Subculturessub=26 Then na$ = "LGBTQ+ fiction" & " " & sd & "2 homoit" & " " & sd & "0 https://homosaurus.org/v3/homoit0000443"
If quick.Subculturessub=27 Then na$ = "LGBTQ+ poetry" & " " & sd & "2 homoit" & " " & sd & "0 https://homosaurus.org/v3/homoit0001118"
If quick.Subculturessub=28 Then na$ = "LGBTQ+ youth literature" & " " & sd & "2 homoit" & " " & sd & "0 https://homosaurus.org/v3/homoit0001522"
If quick.Subculturessub=29 Then na$ = "Queer literature" & " " & sd & "2 homoit" & " " & sd & "0 https://homosaurus.org/v3/homoit0001183"
If quick.Subculturessub=30 Then na$ = "Transgender literature" & " " & sd & "2 homoit" & " " & sd & "0 https://homosaurus.org/v3/homoit0001419"
If quick.Subculturessub=31 Then na$ = "Queer culture" & " " & sd & "2 homoit" & " " & sd & "0 https://homosaurus.org/v3/homoit0001832"
If quick.Subculturessub=32 Then na$ = "Transgender culture" & " " & sd & "2 homoit" & " " & sd & "0 https://homosaurus.org/v3/homoit0001406"
If quick.Subculturessub=33 Then na$ = "Transgender slang" & " " & sd & "2 homoit" & " " & sd & "0 https://homosaurus.org/v3/homoit0001445"
If quick.Subculturessub=34 Then na$ = "Rubber subculture" & " " & sd & "2 homoit" & " " & sd & "0 https://homosaurus.org/v3/homoit0001247"
If quick.Subculturessub=35 Then na$ = "Rubber sex" & " " & sd & "2 homoit" & " " & sd & "0 https://homosaurus.org/v3/homoit0001246"
If quick.Subculturessub=36 Then na$ = "Rubbermen" & " " & sd & "2 homoit" & " " & sd & "0 https://homosaurus.org/v3/homoit0001248"
If quick.Subculturessub=37 Then na$ = "Rubberwomen" & " " & sd & "2 homoit" & " " & sd & "0 https://homosaurus.org/v3/homoit0001249"
If na$ = " " then
else
CS.AddField 1, "650 7" & na$
end if

If quick.LGBTQplus_awardssub=0 Then fa$ = " "
If quick.LGBTQplus_awardssub=1 Then fa$ = "LGBTQ+ awards" & " " & sd & "2 homoit" & " " & sd & "0 https://homosaurus.org/v3/homoit0000818"
If quick.LGBTQplus_awardssub=2 Then fa$ = "LGBTQ+ literary awards" & " " & sd & "2 homoit" & " " & sd & "0 https://homosaurus.org/v3/homoit0000891"
If fa$ = " " then
else
CS.AddField 1, "650 7" & fa$ 
end if

If quick.Body_adornmentsub=0 Then aa$ = " "
If quick.Body_adornmentsub=1 Then aa$ = "Body adornment" & " " & sd & "2 homoit" & " " & sd & "0 https://homosaurus.org/v3/homoit0000214"
If quick.Body_adornmentsub=2 Then aa$ = "Piercings" & " " & sd & "2 homoit" & " " & sd & "0 https://homosaurus.org/v3/homoit0001114"
If quick.Body_adornmentsub=3 Then aa$ = "Tattoos" & " " & sd & "2 homoit" & " " & sd & "0 https://homosaurus.org/v3/homoit0001356"

If aa$ = " " then 
else
CS.AddField 1, "650 7" & aa$ 
end if

If quick.Clothingsub=0 Then ba$ = " "
If quick.Clothingsub=1 Then ba$ = "Clothing" & " " & sd & "2 homoit" & " " & sd & "0 https://homosaurus.org/v3/homoit0000289"
If quick.Clothingsub=2 Then ba$ = "Costumes" & " " & sd & "2 homoit" & " " & sd & "0 https://homosaurus.org/v3/homoit0001808"
If quick.Clothingsub=3 Then ba$ = "Lingerie" & " " & sd & "2 homoit" & " " & sd & "0 https://homosaurus.org/v3/homoit0000970"

If ba$ = " " then
else
CS.AddField 1, "650 7" & ba$
end if

If quick.LGBTQplus_eventssub=0 Then ca$ = " "
If quick.LGBTQplus_eventssub=1 Then ca$ = "LGBTQ+ events" & " " & sd & "2 homoit" & " " & sd & "0 https://homosaurus.org/v3/homoit0000861"
If quick.LGBTQplus_eventssub=2 Then ca$ = "Drag balls" & " " & sd & "2 homoit" & " " & sd & "0 https://homosaurus.org/v3/homoit0000368"
If quick.LGBTQplus_eventssub=3 Then ca$ = "Dyke marches" & " " & sd & "2 homoit" & " " & sd & "0 https://homosaurus.org/v3/homoit0000383"
If quick.LGBTQplus_eventssub=4 Then ca$ = "Gay pride" & " " & sd & "2 homoit" & " " & sd & "0 https://homosaurus.org/v3/homoit0000522"
If quick.LGBTQplus_eventssub=5 Then ca$ = "Gay pride week" & " " & sd & "2 homoit" & " " & sd & "0 https://homosaurus.org/v3/homoit0000523"
If quick.LGBTQplus_eventssub=6 Then ca$ = "Halloween (LGBTQ+ culture)" & " " & sd & "2 homoit" & " " & sd & "0 https://homosaurus.org/v3/homoit0000607"
If quick.LGBTQplus_eventssub=7 Then ca$ = "Lesbian pride" & " " & sd & "2 homoit" & " " & sd & "0 https://homosaurus.org/v3/homoit0000767"
If quick.LGBTQplus_eventssub=8 Then ca$ = "LGBTQ+ sporting events" & " " & sd & "2 homoit" & " " & sd & "0 https://homosaurus.org/v3/homoit0000949"
If quick.LGBTQplus_eventssub=9 Then ca$ = "Trans marches" & " " & sd & "2 homoit" & " " & sd & "0 https://homosaurus.org/v3/homoit0001379"
If quick.LGBTQplus_eventssub=10 Then ca$ = "Women's music festivals" & " " & sd & "2 homoit" & " " & sd & "0 https://homosaurus.org/v3/homoit0001515"
If ca$ = " " then
else
CS.AddField 1, "650 7" & ca$ 
end if

If quick.LGBTQplus_slangsub=0 Then da$ = " "
If quick.LGBTQplus_slangsub=1 Then da$ = "LGBTQ+ slang" & " " & sd & "2 homoit" & " " & sd & "0 https://homosaurus.org/v3/homoit0001317"
If quick.LGBTQplus_slangsub=2 Then da$ = "Bisexual slang" & " " & sd & "2 homoit" & " " & sd & "0 https://homosaurus.org/v3/homoit0000185"
If quick.LGBTQplus_slangsub=3 Then da$ = "Gay slang" & " " & sd & "2 homoit" & " " & sd & "0 https://homosaurus.org/v3/homoit0000535"
If quick.LGBTQplus_slangsub=4 Then da$ = "Lesbian slang" & " " & sd & "2 homoit" & " " & sd & "0 https://homosaurus.org/v3/homoit0000779"
If quick.LGBTQplus_slangsub=5 Then da$ = "Poppers" & " " & sd & "2 homoit" & " " & sd & "0 https://homosaurus.org/v3/homoit0001125"
If quick.LGBTQplus_slangsub=6 Then da$ = "Queer slang" & " " & sd & "2 homoit" & " " & sd & "0 https://homosaurus.org/v3/homoit0001208"
If quick.LGBTQplus_slangsub=7 Then da$ = "Reclaimed terms" & " " & sd & "2 homoit" & " " & sd & "0 https://homosaurus.org/v3/homoit0001559"
If quick.LGBTQplus_slangsub=8 Then da$ = "Crip" & " " & sd & "2 homoit" & " " & sd & "0 https://homosaurus.org/v3/homoit0001891"
If quick.LGBTQplus_slangsub=9 Then da$ = "Dykes" & " " & sd & "2 homoit" & " " & sd & "0 https://homosaurus.org/v3/homoit0000384"
If quick.LGBTQplus_slangsub=10 Then da$ = "Faggots" & " " & sd & "2 homoit" & " " & sd & "0 https://homosaurus.org/v3/homoit0000418"
If quick.LGBTQplus_slangsub=11 Then da$ = "Queer people" & " " & sd & "2 homoit" & " " & sd & "0 https://homosaurus.org/v3/homoit0001195"
If quick.LGBTQplus_slangsub=12 Then da$ = "Transgender slang" & " " & sd & "2 homoit" & " " & sd & "0 https://homosaurus.org/v3/homoit0001445"
If da$ = " " then
else
CS.AddField 1, "650 7" & da$
end if

If quick.LGBTQplus_symbolssub=0 Then ea$ = " "
If quick.LGBTQplus_symbolssub=1 Then ea$ = "LGBTQ+ symbols" & " " & sd & "2 homoit" & " " & sd & "0 https://homosaurus.org/v3/homoit0000954"
If quick.LGBTQplus_symbolssub=2 Then ea$ = "Freedom rings" & " " & sd & "2 homoit" & " " & sd & "0 https://homosaurus.org/v3/homoit0000453"
If quick.LGBTQplus_symbolssub=3 Then ea$ = "Handkerchief codes" & " " & sd & "2 homoit" & " " & sd & "0 https://homosaurus.org/v3/homoit0000609"
If quick.LGBTQplus_symbolssub=4 Then ea$ = "Labrys" & " " & sd & "2 homoit" & " " & sd & "0 https://homosaurus.org/v3/homoit0000696"
If quick.LGBTQplus_symbolssub=5 Then ea$ = "Lambda" & " " & sd & "2 homoit" & " " & sd & "0 https://homosaurus.org/v3/homoit0000698"
If quick.LGBTQplus_symbolssub=6 Then ea$ = "Pink triangles" & " " & sd & "2 homoit" & " " & sd & "0 https://homosaurus.org/v3/homoit0001116"
If quick.LGBTQplus_symbolssub=7 Then ea$ = "Rainbow flags" & " " & sd & "2 homoit" & " " & sd & "0 https://homosaurus.org/v3/homoit0001232"
If quick.LGBTQplus_symbolssub=8 Then ea$ = "Red ribbons (AIDS)" & " " & sd & "2 homoit" & " " & sd & "0 https://homosaurus.org/v3/homoit0000034"

If ea$ = " " then
else
CS.AddField 1, "650 7" & ea$
end if


If quick.LGBTQplus_clubssub=0 Then ga$ = " "
If quick.LGBTQplus_clubssub=1 Then ga$ = "LGBTQ+ clubs" & " " & sd & "2 homoit" & " " & sd & "0 https://homosaurus.org/v3/homoit0000836"
If quick.LGBTQplus_clubssub=2 Then ga$ = "LGBTQ+ book clubs" & " " & sd & "2 homoit" & " " & sd & "0 https://homosaurus.org/v3/homoit0000826"
If quick.LGBTQplus_clubssub=3 Then ga$ = "LGBTQ+ motorcycle clubs" & " " & sd & "2 homoit" & " " & sd & "0 https://homosaurus.org/v3/homoit0000906"
If quick.LGBTQplus_clubssub=4 Then ga$ = "LGBTQ+ sports clubs" & " " & sd & "2 homoit" & " " & sd & "0 https://homosaurus.org/v3/homoit0000950"
If quick.LGBTQplus_clubssub=5 Then ga$ = "Gay sports clubs" & " " & sd & "2 homoit" & " " & sd & "0 https://homosaurus.org/v3/homoit0000537"
If quick.LGBTQplus_clubssub=6 Then ga$ = "Lesbian sports clubs" & " " & sd & "2 homoit" & " " & sd & "0 https://homosaurus.org/v3/homoit0000780"
If quick.LGBTQplus_clubssub=7 Then ga$ = "LGBTQ+ students' clubs" & " " & sd & "2 homoit" & " " & sd & "0 https://homosaurus.org/v3/homoit0000951"
If quick.LGBTQplus_clubssub=8 Then ga$ = "Sex clubs" & " " & sd & "2 homoit" & " " & sd & "0 https://homosaurus.org/v3/homoit0001270"
If ga$ = " " then
else
CS.AddField 1, "650 7" & ga$
end if

If quick.LGBTQplus_night_lifesub=0 Then ja$ = " "
If quick.LGBTQplus_night_lifesub=1 Then ja$ = "LGBTQ+ night life" & " " & sd & "2 homoit" & " " & sd & "0 https://homosaurus.org/v3/homoit0000909"
If quick.LGBTQplus_night_lifesub=2 Then ja$ = "Gay discos" & " " & sd & "2 homoit" & " " & sd & "0 https://homosaurus.org/v3/homoit0000488"
If quick.LGBTQplus_night_lifesub=3 Then ja$ = "Gay saunas" & " " & sd & "2 homoit" & " " & sd & "0 https://homosaurus.org/v3/homoit0000475"
If quick.LGBTQplus_night_lifesub=4 Then ja$ = "Glory holes" & " " & sd & "2 homoit" & " " & sd & "0 https://homosaurus.org/v3/homoit0000595"
If ja$ = " " then
else
CS.AddField 1, "650 7" & ja$
end if

If quick.Miscellanious_culturesub=0 Then ha$ = " "
If quick.Miscellanious_culturesub=1 Then ha$ = "Camp (Gay culture)" & " " & sd & "2 homoit" & " " & sd & "0 https://homosaurus.org/v3/homoit0000238"
If quick.Miscellanious_culturesub=2 Then ha$ = "Gay sensibility" & " " & sd & "2 homoit" & " " & sd & "0 https://homosaurus.org/v3/homoit0000533"
If quick.Miscellanious_culturesub=3 Then ha$ = "Gay theater groups" & " " & sd & "2 homoit" & " " & sd & "0 https://homosaurus.org/v3/homoit0000547"
If quick.Miscellanious_culturesub=4 Then ha$ = "Gaydar" & " " & sd & "2 homoit" & " " & sd & "0 https://homosaurus.org/v3/homoit0000559"
If quick.Miscellanious_culturesub=5 Then ha$ = "Lesbian sensibility" & " " & sd & "2 homoit" & " " & sd & "0 https://homosaurus.org/v3/homoit0000777"
If quick.Miscellanious_culturesub=6 Then ha$ = "Lesbian theater groups" & " " & sd & "2 homoit" & " " & sd & "0 https://homosaurus.org/v3/homoit0000789"
If quick.Miscellanious_culturesub=7 Then ha$ = "LGBTQ+ death notices" & " " & sd & "2 homoit" & " " & sd & "0 https://homosaurus.org/v3/homoit0000849"
If quick.Miscellanious_culturesub=8 Then ha$ = "LGBTQ+ film festivals" & " " & sd & "2 homoit" & " " & sd & "0 https://homosaurus.org/v3/homoit0000868"
If quick.Miscellanious_culturesub=9 Then ha$ = "LGBTQ+ monuments" & " " & sd & "2 homoit" & " " & sd & "0 https://homosaurus.org/v3/homoit0000642"
If quick.Miscellanious_culturesub=10 Then ha$ = "LGBTQ+ older people's organisations" & " " & sd & "2 homoit" & " " & sd & "0 https://homosaurus.org/v3/homoit0000856"
If quick.Miscellanious_culturesub=11 Then ha$ = "LGBTQ+ weddings" & " " & sd & "2 homoit" & " " & sd & "0 https://homosaurus.org/v3/homoit0000963"
If quick.Miscellanious_culturesub=12 Then ha$ = "Musicals" & " " & sd & "2 homoit" & " " & sd & "0 https://homosaurus.org/v3/homoit0001031"
If quick.Miscellanious_culturesub=13 Then ha$ = "Nudism" & " " & sd & "2 homoit" & " " & sd & "0 https://homosaurus.org/v3/homoit0001038"
If quick.Miscellanious_culturesub=14 Then ha$ = "Sex scandals" & " " & sd & "2 homoit" & " " & sd & "0 https://homosaurus.org/v3/homoit0001273"
If quick.Miscellanious_culturesub=15 Then ha$ = "Social privilege" & " " & sd & "2 homoit" & " " & sd & "0 https://homosaurus.org/v3/homoit0001133"
If quick.Miscellanious_culturesub=16 Then ha$ = "Voguing" & " " & sd & "2 homoit" & " " & sd & "0 https://homosaurus.org/v3/homoit0001494"
If ha$ = " " then 
else
CS.AddField 1, "650 7" & ha$ 
end if

Cancelled:   
End Sub

