Imports System.Data
Imports System.Data.SqlClient
Imports iTextSharp.text
Imports iTextSharp.text.pdf
Imports iTextSharp.text.pdf.parser
Imports System.IO
Public Class Form1
    Dim IMG As Image
    Dim con As SqlConnection = New SqlConnection("Data Source=COMP2\SQLEXPRESS;Initial Catalog=thobs2021;Integrated Security=True")
    Dim cmd As SqlCommand = New SqlCommand()
    Dim sda As SqlDataAdapter = New SqlDataAdapter()
    Dim dt As DataTable = New DataTable()
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        cmd = New SqlCommand("select OSCH_NO,OBSRNAME,OADD1,OADD2,OADD3,OADD4,OADD5,OPIN,CEN_NO,CABBR_NAME,cs_schno,CS_SCHNAME,cpmobile,DATE1,CLS1,SUB1,SUBNAME1,DATE2,CLS2,SUB2,SUBNAME2,DATE3,CLS3,SUB3,SUBNAME3,DATE4,CLS4,SUB4,SUBNAME4,DATE5,CLS5,SUB5,SUBNAME5,DATE6,CLS6,SUB6,SUBNAME6,DATE7,CLS7,SUB7,SUBNAME7,DATE8,CLS8,SUB8,SUBNAME8,DATE9,CLS9,SUB9,SUBNAME9,DATE10,CLS10,SUB10,SUBNAME10,DATE11,CLS11,SUB11,SUBNAME11,DATE12,CLS12,SUB12,SUBNAME12,DATE13,CLS13,SUB13,SUBNAME13,DATE14,CLS14,SUB14,SUBNAME14,DATE15,CLS15,SUB15,SUBNAME15,DATE16,CLS16,SUB16,SUBNAME16,DATE17,CLS17,SUB17,SUBNAME17,DATE18,CLS18,SUB18,SUBNAME18,DATE19,CLS19,SUB19,SUBNAME19,DATE20,CLS20,SUB20,SUBNAME20,DATE21,CLS21,SUB21,SUBNAME21,DATE22,CLS22,SUB22,SUBNAME22,DATE23,CLS23,SUB23,SUBNAME23,DATE24,CLS24,SUB24,SUBNAME24,DATE25,CLS25,SUB25,SUBNAME25,omobile,slno,POST,cadd1,cadd2,cadd3,cadd4,cadd5,cpin FROM thobs2022  ", con)
        'cmd = New SqlCommand("select OSCH_NO,OBSRNAME,OADD1,OADD2,OADD3,OADD4,OADD5,OPIN,CEN_NO,CABBR_NAME,cs_schno,CS_SCHNAME,cpmobile,DATE1,CLS1,SUB1,SUBNAME1,DATE2,CLS2,SUB2,SUBNAME2,DATE3,CLS3,SUB3,SUBNAME3,DATE4,CLS4,SUB4,SUBNAME4,DATE5,CLS5,SUB5,SUBNAME5,DATE6,CLS6,SUB6,SUBNAME6,DATE7,CLS7,SUB7,SUBNAME7,DATE8,CLS8,SUB8,SUBNAME8,DATE9,CLS9,SUB9,SUBNAME9,DATE10,CLS10,SUB10,SUBNAME10  FROM theory_observer  ", con)

        sda = New SqlDataAdapter(cmd)

        Dim dta As DataTable = New DataTable()
        sda.Fill(dta)
        Dim cntr As Integer = dta.Rows.Count
        For j = 0 To cntr - 1
            Dim doc As New Document(PageSize.A4, 30, 30, 50, 50)
            Dim pw = PdfWriter.GetInstance(doc, New FileStream("e:\thob\" + dta.Rows(j)(0) + "_" + dta.Rows(j)(114) + "_" + dta.Rows(j)(8) + "_thobs.pdf", FileMode.Create))
            doc.Open()
            IMG = Image.GetInstance("C:\Users\HP\source\repos\pdf_letter\HEADER.PNG")
            IMG.ScaleToFit(525.0F, 200.0F)
            IMG.SetAbsolutePosition(15, 750)
            doc.Add(IMG)
            IMG = Image.GetInstance("C:\Users\HP\source\repos\pdf_letter\FOOTER.PNG")
            IMG.ScaleToFit(525.0F, 200.0F)
            IMG.SetAbsolutePosition(20, 20)
            doc.Add(IMG)
            Dim p1 As New Paragraph("

============================================================================
NO.CBSE/RO(PTN)/CONF/OBSERVER/2022/Term 2                                 DATED : " + Module1.dt + "
                                                                          Observer No - " + dta.Rows(j)(114).ToString + "                                              
    " + dta.Rows(j)(1).ToString + " < " + dta.Rows(j)(0).ToString + "   >               
    " + dta.Rows(j)(2).ToString + "                  
    " + dta.Rows(j)(3).ToString + "                     
    " + dta.Rows(j)(4).ToString + "   
    " + dta.Rows(j)(5).ToString + " 
    " + dta.Rows(j)(6).ToString + " ,  PIN - " + dta.Rows(j)(7).ToString + "             

SUB : APPOINTMENT OF THEORY OBSERVER FOR Term 2 AISSE/AISSCE (MAIN)-2022.
")

            p1.Alignment = Element.ALIGN_JUSTIFIED
            doc.Add(p1)
            Dim p2 As New Paragraph("
SIR / MADAM,

    I am to inform you that the Board is  pleased  to appoint you as an observer for the Secondary/Senior School Term 2 Main Examination 2022 commencing from 26th April 2022  at  the  below  mentioned  centre  on  the  specified dates.
    Therefore, You are requested to kindly make  it  convenient to inspect the following examination centre(s) during Examinations.

")
            p2.Alignment = Element.ALIGN_JUSTIFIED
            doc.Add(p2)
            Dim table2 As New PdfPTable(3)
            table2.TotalWidth = 480.0F
            table2.LockedWidth = True
            table2.DefaultCell.Border = 0
            table2.DefaultCell.Left = 0


            table2.HorizontalAlignment = 0
            Dim widths2 As Single() = New Single() {3.0F, 10.0F, 5.0F}
            table2.SetWidths(widths2)
            Dim arial2 As BaseFont = BaseFont.CreateFont("c:\windows\fonts\arial.ttf", BaseFont.IDENTITY_H, BaseFont.EMBEDDED)
            Dim font2 As New Font(arial2, 13, 1)
            table2.AddCell("Centre No.")
            ' table2.AddCell(New Phrase("Sl No.", font2))
            table2.AddCell("Name Of the Centre")
            table2.AddCell("Contact No")
            table2.AddCell(dta.Rows(j)(8).ToString)
            table2.AddCell(dta.Rows(j)(9).ToString)
            table2.AddCell(dta.Rows(j)(12).ToString)
            doc.Add(table2)
            Dim p2_ As New Paragraph("
Date(s) of Inspection :

")
            doc.Add(p2_)
            'p2_.Alignment = Element.ALIGN_JUSTIFIED
            Dim table3 As New PdfPTable(1)
            table3.TotalWidth = 100.0F
            table3.LockedWidth = True
            ' table3.DefaultCell.Border = 1
            table3.DefaultCell.Left = 0


            table3.HorizontalAlignment = 0
            Dim widths3 As Single() = New Single() {3.0F}
            'Dim widths3 As Single() = New Single() {3.0F, 2.0F, 2.0F, 7.0F}
            table3.SetWidths(widths3)
            Dim arial3 As BaseFont = BaseFont.CreateFont("c:\windows\fonts\arial.ttf", BaseFont.IDENTITY_H, BaseFont.EMBEDDED)
            Dim font3 As New Font(arial3, 13, 1)
            table3.AddCell("Date of Exam")
            ''''''''''''''''''''''''''''
            ' Dim a As Integer
            ' Dim k As Integer
            '  Dim k1 As Integer
            '  k = 13
            ' For a = 1 To 24
            'If a = 10 Then
            'k1 = k
            'Exit For
            'End If
            'table3.AddCell(dta.Rows(j)(k).ToString)
            '= k + 1
            'If dta.Rows(j)(k).ToString = "10" Then
            'table3.AddCell("<" + dta.Rows(j)(k + 1).ToString + ">  " + dta.Rows(j)(k + 2).ToString)
            'table3.AddCell("  ")
            'ElseIf dta.Rows(j)(k).ToString = "12" Then
            'table3.AddCell("  ")
            'table3.AddCell("<" + dta.Rows(j)(k + 1).ToString + ">  " + dta.Rows(j)(k + 2).ToString)
            'End If
            'k = k + 3
            'Next
            'doc.Add(table3)
            '    If a = 10 Then
            '        doc.NewPage()
            '        IMG = Image.GetInstance("C:\Users\HP\source\repos\pdf_letter\HEADER.PNG")
            '        IMG.ScaleToFit(525.0F, 200.0F)
            '        IMG.SetAbsolutePosition(15, 750)
            '        doc.Add(IMG)
            '        IMG = Image.GetInstance("C:\Users\HP\source\repos\pdf_letter\FOOTER.PNG")
            '        IMG.ScaleToFit(525.0F, 200.0F)
            '        IMG.SetAbsolutePosition(20, 20)
            '        doc.Add(IMG)
            '        doc.Add(IMG)
            '        Dim p256 As New Paragraph("
            '

            '")
            '    doc.Add(p256)
            '    Dim table4 As New PdfPTable(3)
            '    table4.TotalWidth = 540.0F
            '    table4.LockedWidth = True
            '    table4.DefaultCell.Border = 0
            '    table4.DefaultCell.Left = 0


            'table4.HorizontalAlignment = 0
            '    Dim widths4 As Single() = New Single() {2.0F, 5.0F, 5.0F}
            '    table4.SetWidths(widths4)
            '    Dim arial4 As BaseFont = BaseFont.CreateFont("c:\windows\fonts\arial.ttf", BaseFont.IDENTITY_H, BaseFont.EMBEDDED)
            '    Dim font4 As New Font(arial4, 13, 1)
            '    table4.AddCell("Date of Exam")
            '    table4.AddCell("Class X")
            '    table4.AddCell("Class XII")
            '    k = k1
            '    For a = 10 To 24
            '
            '        table4.AddCell(dta.Rows(j)(k).ToString)
            '        k = k + 1
            '        If dta.Rows(j)(k).ToString = "10" Then
            '            table4.AddCell("<" + dta.Rows(j)(k + 1).ToString + ">  " + dta.Rows(j)(k + 2).ToString)
            '            table4.AddCell("  ")
            '        ElseIf dta.Rows(j)(k).ToString = "12" Then
            '            table4.AddCell("  ")
            '            table4.AddCell("<" + dta.Rows(j)(k + 1).ToString + ">  " + dta.Rows(j)(k + 2).ToString)
            '        End If
            '        k = k + 3
            '    Next
            '    doc.Add(table4)
            'End If
            ''''''''''''''''''''''''''''
            Dim a As Integer
            Dim k As Integer
            k = 13
            For a = 1 To 113
                If String.IsNullOrEmpty(dta.Rows(j)(k).ToString) Then
                    Exit For
                End If
                table3.AddCell(dta.Rows(j)(k).ToString)
                k = k + 1
                'table3.AddCell(dta.Rows(j)(k).ToString)
                'table3.AddCell(dta.Rows(j)(k + 1).ToString)
                'table3.AddCell(dta.Rows(j)(k + 2).ToString)
                k = k + 3
            Next
            doc.Add(table3)
            doc.NewPage()
            IMG = Image.GetInstance("C:\Users\HP\source\repos\pdf_letter\HEADER.PNG")
            IMG.ScaleToFit(525.0F, 200.0F)
            IMG.SetAbsolutePosition(15, 750)
            doc.Add(IMG)
            IMG = Image.GetInstance("C:\Users\HP\source\repos\pdf_letter\FOOTER.PNG")
            IMG.ScaleToFit(525.0F, 100.0F)
            IMG.SetAbsolutePosition(10, 10)
            doc.Add(IMG)
            Dim p3 As New Paragraph("

    As an observer you may visit to the said centre(s) on the dates  mentioned above during the Examination and send your Confidential  Report  to  the  Regional  Officer in  the  prescribed   proforma immediately  after  the  inspection  indicating  whether  the  Examination  was  conducted smoothly and whether  all  the  arrangments  made  were  satisfactory (copies of Proforma are enclosed).
")
            p3.Alignment = Element.ALIGN_JUSTIFIED
            doc.Add(p3)
            Dim p4 As New Paragraph("
    While performing the duty as  Observer you are requested to  particularly verify  the  sealed  packets  containing  Question Papers which should  Not  be opened before the prescribed time As  per Date sheet. For this purpose, you may  visit  the  centre  well in advance  before  commencement  of  the  Examination  And  verify  the intactness of Question Paper envelops taken from the custodian bank  of  the  exam  centres  by the Centre Superintendent And record the facts On the sealed envelop. In case you find any difference in  the  seal impression etc.on the envelops, the matter may immediately be brought to the notice of the undersigned.You  may  kindly  stay  at  the  centre  till  completion  of  the Examination of the day And also ensure that the packing of Answer Books Is done in your presence.")
            p4.Alignment = Element.ALIGN_JUSTIFIED
            doc.Add(p4)
            Dim p5 As New Paragraph("
    FURTHER THE OBSERVERS ARE PERSONALLY REQUIRED TO VERIFY  THAT AFTER THE EXAMINATION IS OVER, SEALED ANSWER SCRIPTS PARCELS AND RELATED MATERIALS OF THE DAY, ARE STRICTLY  DISPATCHED  ON  THE SAME  DAY  TO THE  REGIONAL OFFICE, CBSE PATNA BY INSURED SPEED POST  IN  SEPARATE  SEALED  COVER  OR  HANDED  OVER   TO  THE OFFECER/OFFICIAL DEPUTED FROM THIS OFFICE.  IN  NO  CASE  IT IS  TO BE KEPT AT THE  CENTRE SCHOOL  OVERNIGHT  AND THESE  ARE  ALSO  NOT  TO  BE  DESPATCHED THROUGH RAILWAY PARCEL SERVICE.")
            p5.Alignment = Element.ALIGN_JUSTIFIED
            doc.Add(p5)
            Dim p6 As New Paragraph("
    THE FOLLOWING POINTS MUST BE CHECKED/ENSURED BY THE OBSERVER AT EXAMINATION CENTRE:- ")
            p6.Alignment = Element.ALIGN_JUSTIFIED
            doc.Add(p6)
            Dim p7 As New Paragraph("
1)  THE CENTRE SUPDT. IS STRICTLY INSTRUCTED THAT HE SHOULD WRITE THE ADDRESS ON THE  SEALED  ANSWER  BOOK  PARCELS  OF  CLASS X WITH  RED  COLOR  AND ON THE PARCELS   OF  CLASS  XII  WITH  BLUE  COLOR   SO  THAT  IT   BECOMES  EASILY DISTINGUISHABLE WHETHER THE PARCEL CONTAINS THE ANSWER BOOKS ARE OF CLASS X/XII, HOWEVER CS GUIDELINES FOR AISSE/AISSCE - Term 2 2022 WILL BE PROVIDED  TO YOU IN DUE COURSE OF TIME.")
            p7.Alignment = Element.ALIGN_JUSTIFIED
            doc.Add(p7)
            Dim p8 As New Paragraph("
2)   ON ALL OTHER PARCELS CONTAINING  MATERIAL  OTHER  THAN  ANSWER  BOOKS OF THE EXAMINATION, NOT RELATED  TO  CSO  (SECRECY WORK), THE  CENTRE  SUPDT. SHALL WRITE THE ADDRESS  IN  BLUE  COLOR BUT IN  THE  BOTTOM  HE MUST WRITE WITHIN  BRACKET 'NOT  FOR  CSO' SO THAT  THIS  MATERIAL  COULD  BE  IMMEDIATELY")
            p8.Alignment = Element.ALIGN_JUSTIFIED
            doc.Add(p8)
            doc.NewPage()
            IMG = Image.GetInstance("C:\Users\HP\source\repos\pdf_letter\HEADER.PNG")
            IMG.ScaleToFit(525.0F, 200.0F)
            IMG.SetAbsolutePosition(15, 750)
            doc.Add(IMG)
            IMG = Image.GetInstance("C:\Users\HP\source\repos\pdf_letter\FOOTER.PNG")
            IMG.ScaleToFit(525.0F, 200.0F)
            IMG.SetAbsolutePosition(20, 20)
            doc.Add(IMG)
            Dim p9 As New Paragraph(" 

 BE  SEGREGATED ACCORDINGLY IN  THE REGIONAL  OFFICE  AND MAY NOT BE HANDED OVER  TO CSOs (SECRECY TEAM).
")
            p9.Alignment = Element.ALIGN_JUSTIFIED
            doc.Add(p9)
            Dim p10 As New Paragraph(" 
3)  SUBJECT  CODE, NAME OF THE  SUBJECT, CENTRE NO., CENTRE  NAME  AND  DATE  OF EXAMINATION MUST BE MENTIONED  ON THE ANSWER BOOKS PARCELS CLEARLY  AND  THE  ANSWER  BOOKS OF EACH SUBJECT SHOULD BE PACKED SEPRATELY.")
            p10.Alignment = Element.ALIGN_JUSTIFIED
            doc.Add(p10)
            Dim p11 As New Paragraph(" 
4)  THE ANSWER BOOKS PERTAINING TO  UNFAIRMEANS CASES  AND PHYSICALLY CHALLANGED (SPASTIC, BLIND, PHYSICALLY  HANDICAPPED  AND  DYSLEXEC  CHILDREN) SHOULD BE  PACKED  IN A  SEPRATE SEALED  ENVELOP  SENT  BY  THE BOARD  AND SAME WILL BE DESPATCHED BY SPEED POST SEPERATELY.")

            p11.Alignment = Element.ALIGN_JUSTIFIED
            doc.Add(p11)
            Dim p12 As New Paragraph(" 
5)   DURING  PREVIOUS YEAR EXAMINATION IT  HAS  BEEN  NOTICED IN  SOME CASES THAT THE PRESENT CANDIDATES HAVE BEEN MARKED ABSENT WHILE  ABSENT CANDIDATES HAVE  BEEN MARKED PRESENT. ALSO AT SOME PLACES SIGNATURE OF THE CANDIDATE  WAS NOT TAKEN IN THE  ATTANDANCE SHEET WHICH MUST BE  VERIFIED.  CENTRE  SUPDTS  MAY PLEASE BE INSTRUCTED FOR NON-REPEATATION OF SUCH MISTAKES.")
            p12.Alignment = Element.ALIGN_JUSTIFIED
            doc.Add(p12)
            Dim p13 As New Paragraph("
6)   A COPY OF GUIDELINES FOR CENTRE SUPDTS AND  ASSTT. SUPDTS. TERM 2 2022  HAS ALREADY BEEN PROVIDED TO THE CENTRE SUPDT. FOR STRICT COMPLIANCE AND  REFERENCES. IN CASE OBSERVER FEELS ANY DIFFICULTY/PROBLEM/DEVIATION AT THE CENTRE CONCERNED, THE SAME MAY BE PERUSED. ")
            p13.Alignment = Element.ALIGN_JUSTIFIED
            doc.Add(p13)
            Dim p14 As New Paragraph("
    Your observation/ report regarding conduct  of the  Examination at  the centre(s) may be submitted to  the Regional Officer immediately after  visit at the Examination Centre(s) for onward information to the Competant Authority  of the Board.")
            p14.Alignment = Element.ALIGN_JUSTIFIED
            doc.Add(p14)
            Dim p15 As New Paragraph("
    For performing the assignment as observer a  remuneration  of  Rs. 500/- per Day will be paid to you besides conveyance charges @ Rs. 250/- ( for local observers )  per day  And  T.A./D.A. ( for outside observers )  Is  permissible as per Boards rules for performing the  assignment to out station centres.")
            p15.Alignment = Element.ALIGN_JUSTIFIED
            doc.Add(p15)
            doc.NewPage()
            IMG = Image.GetInstance("C:\Users\HP\source\repos\pdf_letter\HEADER.PNG")
            IMG.ScaleToFit(525.0F, 200.0F)
            IMG.SetAbsolutePosition(15, 750)
            doc.Add(IMG)
            IMG = Image.GetInstance("C:\Users\HP\source\repos\pdf_letter\FOOTER.PNG")
            IMG.ScaleToFit(525.0F, 200.0F)
            IMG.SetAbsolutePosition(20, 20)
            doc.Add(IMG)
            Dim p16 As New Paragraph("

    YOU ARE REQUESTED TO SEND THE ACCEPTANCE OF  THE ASSIGNMENT  OFFERED BY THE BOARD  IN  ENCLOSED  PROFORMA DULY COMPLETED IN  ALL RESPECT ALONGWITH  THE CERTFICATION IN REGARD WITH NON APPEARING OF NEAR RELATION AT AISSE/AISSCE TERM 2 2022 BY RETURN OF POST/ BY EMAIL (abcell.ropatna@cbseshiksha.in). SO, AS TO REACH THE UNDERSIGNED IMMEIDATELY.

    In case your near relation is appearing in the said exam, the  same may be informed immediately,  and  offer  of  the  said assignment may  treated  as cancelled.")
            p16.Alignment = Element.ALIGN_JUSTIFIED
            doc.Add(p16)
            Dim p17 As New Paragraph("
YOURS FAITHFULLY")
            p17.Alignment = Element.ALIGN_RIGHT
            doc.Add(p17)
            IMG = Image.GetInstance("C:\Users\HP\source\repos\pdf_letter\rosignpng.PNG")
            IMG.ScaleToFit(80.0F, 80.0F)
            IMG.SetAbsolutePosition(470, 545)
            doc.Add(IMG)
            Dim p18 As New Paragraph(" 

(JAGADISH BARMAN)
 REGIONAL OFFICER
   CBSE, PATNA

")
            p18.Alignment = Element.ALIGN_RIGHT
            doc.Add(p18)
            Dim P50 As New Paragraph("
Copy to: Centre Superintendent  " + dta.Rows(j)(8).ToString + " /  " + dta.Rows(j)(10).ToString + " 
" + dta.Rows(j)(116).ToString + " 
" + dta.Rows(j)(117).ToString + "
" + dta.Rows(j)(118).ToString + "
" + dta.Rows(j)(119).ToString + "
" + dta.Rows(j)(120).ToString + " PIN: " + dta.Rows(j)(121).ToString + "
ADDRESS FOR INFORMATION")
            ' P50.Alignment = Element.ALIGN_JUSTIFIED
            doc.Add(P50)
            Dim p19 As New Paragraph(" 













Encl: As above.")
            p19.Alignment = Element.ALIGN_JUSTIFIED
            doc.Add(p19)
            doc.NewPage()
            IMG = Image.GetInstance("C:\Users\HP\source\repos\pdf_letter\acceptance_header.PNG")
            IMG.ScaleToFit(525.0F, 200.0F)
            IMG.SetAbsolutePosition(15, 750)
            doc.Add(IMG)
            IMG = Image.GetInstance("C:\Users\HP\source\repos\pdf_letter\FOOTER.PNG")
            IMG.ScaleToFit(525.0F, 200.0F)
            IMG.SetAbsolutePosition(20, 20)
            doc.Add(IMG)
            Dim p20 As New Paragraph("


ACCEPTANCE FORM OF THEORY OBSERVER FOR AISSE/AISSCE TERM 2 2022
        -------------------------------------------------------------------------------------------------------------------
            (To be sent by Speed Post/Fax or through Messanger)

IMMEDIATE & CONFIDENTIAL")
            p20.Alignment = Element.ALIGN_CENTER

            doc.Add(p20)
            Dim p21 As New Paragraph("                                                                                                       Observer No -  " + dta.Rows(j)(114).ToString + "
The Regional Officer 
Central Board of Secondary Education     
Ambika Complex, Behind State Bank Colony, Near Brahmsthan 
Sheikhpura, Bailey Road, Patna- 800014 (Bihar)

Sir,
     With reference to your letter No. RO(PTN)/CONF/OBSERVER/ " + dta.Rows(j)(8).ToString + " /TERM 2/2022 dated    " + Module1.dt + " ,   I  hereby  express    my    willingness  to act as  Observer for AISSE/AISSCE - Term 2 2022. I  shall  do  this  work with  perfect  efficiency and according   to   the instructions issued by the Board.
     
     I  hereby certify that none of my near relation is appearing in the aforesaid Examinations of the Board.

")
            p21.Alignment = Element.ALIGN_JUSTIFIED
            doc.Add(p21)
            Dim p22 As New Paragraph("                                                                                Yours faithfully,")
            'p22.Alignment = Element.ALIGN_RIGHT
            doc.Add(p22)
            Dim p23 As New Paragraph("
                                              Signature (with date):                                                              
                                    ")
            doc.Add(p23)

            Dim table As New PdfPTable(2)
            table.TotalWidth = 500.0F
            table.LockedWidth = True
            table.DefaultCell.Border = 0
            table.DefaultCell.Left = 0


            table.HorizontalAlignment = 2
            Dim widths As Single() = New Single() {6.0F, 7.0F}
            table.SetWidths(widths)
            Dim arial As BaseFont = BaseFont.CreateFont("c:\windows\fonts\arial.ttf", BaseFont.IDENTITY_H, BaseFont.EMBEDDED)
            Dim font As New Font(arial, 13)
            table.AddCell(" ")
            table.AddCell(dta.Rows(j)(0).ToString + "   " + dta.Rows(j)(1).ToString)
            ''
            table.AddCell(" ")
            table.AddCell(dta.Rows(j)(115).ToString)
            ''
            table.AddCell(" ")
            table.AddCell(dta.Rows(j)(2).ToString)
            table.AddCell(" ")
            table.AddCell(dta.Rows(j)(3).ToString)
            table.AddCell(" ")
            table.AddCell(dta.Rows(j)(4).ToString)
            table.AddCell(" ")
            table.AddCell(dta.Rows(j)(5).ToString)
            table.AddCell(" ")
            table.AddCell(dta.Rows(j)(6).ToString + "  ,  " + dta.Rows(j)(7).ToString)
            doc.Add(table)
            Dim p25_ As New Paragraph("                                                         Mobile No. :    " + dta.Rows(j)(113).ToString + "")
            doc.Add(p25_)

            Dim p24 As New Paragraph("





Dated: " + Module1.dt + "")
            p24.Alignment = Element.ALIGN_JUSTIFIED
            doc.Add(p24)
            doc.Close()
        Next
        MessageBox.Show("Voilla! Files Created.")
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        'cmd = New SqlCommand("select OSCH_NO,OBSRNAME,OADD1,OADD2,OADD3,OADD4,OADD5,OPIN,CEN_NO,CABBR_NAME,cs_schno,CS_SCHNAME,cpmobile,DATE1,SUB1,SUBNAME1,CLS1,DATE2,SUB2,SUBNAME2,CLS2,DATE3,SUB3,SUBNAME3,CLS3,DATE4,SUB4,SUBNAME4,CLS4,DATE5,SUB5,SUBNAME5,CLS5,DATE6,SUB6,SUBNAME6,CLS6,DATE7,SUB7,SUBNAME7,CLS7,DATE8,SUB8,SUBNAME8,CLS8,DATE9,SUB9,SUBNAME9,CLS9,DATE10,SUB10,SUBNAME10,CLS10,DATE11,SUB11,SUBNAME11,CLS11,DATE12,SUB12,SUBNAME12,CLS12,DATE13,SUB13,SUBNAME13,CLS13,DATE14,SUB14,SUBNAME14,CLS14,DATE15,SUB15,SUBNAME15,CLS15,DATE16,SUB16,SUBNAME16,CLS16,DATE17,SUB17,SUBNAME17,CLS17,DATE18,SUB18,SUBNAME18,CLS18,DATE19,SUB19,SUBNAME19,CLS19,DATE20,SUB20,SUBNAME20,CLS20,omobile,slno,POST FROM thobs_letter  ", con)
        'sda = New SqlDataAdapter(cmd)
        'Dim dta As DataTable = New DataTable()
        'sda.Fill(dta)
        'Dim cntr As Integer = dta.Rows.Count
        'Dim doc As New Document(PageSize.A4, 30, 30, 50, 50)
        'Dim pw = PdfWriter.GetInstance(doc, New FileStream("e:\thob\" + dta.Rows(0)(0) + "_theory_observer_appointment_letter.pdf", FileMode.Create))
        'doc.Open()
        'IMG = Image.GetInstance("C:\Users\HP\source\repos\pdf_letter\HEADER.PNG")
        'IMG.ScaleToFit(525.0F, 200.0F)
        'IMG.SetAbsolutePosition(15, 750)
        'doc.Add(IMG)
        'IMG = Image.GetInstance("C:\Users\HP\source\repos\pdf_letter\FOOTER.PNG")
        'IMG.ScaleToFit(525.0F, 200.0F)
        'IMG.SetAbsolutePosition(20, 20)
        'doc.Add(IMG)
        'Dim x As Integer
        'For x = 0 To 100
        'Dim p As New Paragraph(x.ToString)
        'doc.Add(p)
        'Next
        'doc.Close()

    End Sub
End Class
