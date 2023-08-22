Imports System.Data.SqlClient

Public Class Form1

    Dim con As SqlConnection = New SqlConnection("Data Source=.\SQLEXPRESS;Initial Catalog=Ujjwal204;Integrated Security=True")
    Private Sub btnCancel_Click(sender As Object, e As EventArgs) Handles btnCancel.Click
        pnlFormLayout.Visible = True
        pnlFormPage.Visible = False
    End Sub

    Private Sub btnNext_Click(sender As Object, e As EventArgs) Handles btnNext.Click
        pnlFormLayout.Visible = False
        pnlFormPage.Visible = True
    End Sub

    Private Sub btnExit_Click(sender As Object, e As EventArgs) Handles btnExit.Click
        Me.Close()
    End Sub
    Private Sub timerForm_Tick(sender As Object, e As EventArgs) Handles timerForm.Tick
        txtFormDate.Text = DateTime.Now.ToString("dd/MM/yyyy")
        txtFormTime.Text = DateTime.Now.ToString("hh:mm:ss")
        timerForm.Start()
    End Sub

    Private Sub dtpDOB_ValueChanged(sender As Object, e As EventArgs) Handles dtpDOB.ValueChanged
        Dim date_DOB As DateTime
        Dim age_span As TimeSpan
        Dim age_years As Integer
        date_DOB = Convert.ToDateTime(dtpDOB.Value)
        age_span = Date.Now.Subtract(date_DOB)
        age_years = Math.Truncate(age_span.TotalDays / 365)
        txtAge.Text = age_years.ToString
    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        'TODO: This line of code loads data into the 'Ujjwal204DataSet2.Language_table' table. You can move, or remove it, as needed.
        Me.Language_tableTableAdapter.Fill(Me.Ujjwal204DataSet2.Language_table)
        'TODO: This line of code loads data into the 'Ujjwal204DataSet1.Branch_tbl' table. You can move, or remove it, as needed.
        Me.Branch_tblTableAdapter.Fill(Me.Ujjwal204DataSet1.Branch_tbl)
        'TODO: This line of code loads data into the 'Ujjwal204DataSet.City_Table' table. You can move, or remove it, as needed.
        Me.City_TableTableAdapter.Fill(Me.Ujjwal204DataSet.City_Table)
        'TODO: This line of code loads data into the 'Ujjwal204DataSet.State_Table' table. You can move, or remove it, as needed.
        Me.State_TableTableAdapter.Fill(Me.Ujjwal204DataSet.State_Table)

        Dim cmd As New SqlCommand("SELECT DISTINCT City_StateCode,StateName FROM BankStateCityView", con)
        Dim da As New SqlDataAdapter(cmd)
        Dim dt As New DataTable
        da.Fill(dt)
        comboxState.DataSource = dt
        comboxState.DisplayMember = "StateName"
        comboxState.ValueMember = "City_StateCode"
        comboxState.Text = "Select State"

    End Sub

    Private Sub txtFormNo_TextChanged(sender As Object, e As EventArgs) Handles txtFormNo.FontChanged
        Dim myCmd As New SqlCommand("SELECT count(*) from Bank_Account_Customer", con)
        con.Open()
        txtFormNo.Text = (CInt(myCmd.ExecuteScalar) + 1).ToString
        con.Close()
    End Sub
    Private Sub btnSave_Click(sender As Object, e As EventArgs) Handles btnSave.Click
        Dim myCmd As New SqlCommand("INSERT INTO Bank_Account_Customer (FormDate,FormTime,Title,FirstName,MiddleName,LastName,Gender,DOB,Age,STDcode,Telephone,MobileNo,email,State,City,BranchName,AccountType,Language) VALUES ('" & Now.ToString("dd MMMM yyyy") & "','" & Now.ToString("hh:mm:ss") & "','" & lstboxTitle.SelectedIndex + 1 & "','" & txtFirstname.Text & "','" & txtMiddlename.Text & "','" & txtLastName.Text & "','" & lstboxSex.SelectedIndex + 1 & "','" & dtpDOB.Text & "','" & CInt(txtAge.Text) & "','" & txtSTD.Text & "','" & txtTelephone.Text & "','" & txtMobileNo.Text & "','" & txtEmail.Text & "','" & comboxState.SelectedValue & "','" & comboxCity.SelectedValue & "','" & comboxBranchName.SelectedValue & "','" & lstboxAccountType.SelectedIndex + 1 & "','" & comboxLanguage.SelectedValue & "')", con)
        con.Open()
        If myCmd.ExecuteNonQuery() = 1 Then
            MessageBox.Show("Custumer Details added")
            Dim myCmdSubmit As New SqlCommand("SELECT count(*) from Bank_Account_Customer", con)
            txtFormNo.Text = (CInt(myCmdSubmit.ExecuteScalar) + 1).ToString
            txtFirstname.Text = ""
            txtMiddlename.Text = ""
            txtLastName.Text = ""
            txtAge.Text = ""
            txtSTD.Text = ""
            txtTelephone.Text = ""
            txtMobileNo.Text = ""
            txtEmail.Text = ""
            comboxState.Text = ""
            comboxCity.Text = ""
            comboxBranchName.Text = ""
            comboxLanguage.Text = ""
        Else
            MessageBox.Show("Details not added")
        End If

        con.Close()

    End Sub

    Private Sub comboxState_SelectionChangeCommitted(sender As Object, e As EventArgs) Handles comboxState.SelectionChangeCommitted
        Dim cmd As New SqlCommand("SELECT DISTINCT CityName,bCityCode FROM BankStateCityView WHERE City_StateCode=@StateCode", con)
        cmd.Parameters.AddWithValue("StateCode", comboxState.SelectedValue)
        Dim da As New SqlDataAdapter(cmd)
        Dim dt As New DataTable
        da.Fill(dt)
        comboxCity.DataSource = dt
        comboxCity.DisplayMember = "CityName"
        comboxCity.ValueMember = "bCityCode"
        comboxCity.Text = "Select City"
    End Sub

    Private Sub comboxCity_SelectionChangeCommitted(sender As Object, e As EventArgs) Handles comboxCity.SelectionChangeCommitted
        Dim cmd As New SqlCommand("SELECT DISTINCT branchName,branchCode FROM BankStateCityView WHERE bCityCode=@CityCode", con)
        cmd.Parameters.AddWithValue("CityCode", comboxCity.SelectedValue)
        Dim da As New SqlDataAdapter(cmd)
        Dim dt As New DataTable
        da.Fill(dt)
        comboxBranchName.DataSource = dt
        comboxBranchName.DisplayMember = "branchName"
        comboxBranchName.ValueMember = "branchCode"
        comboxBranchName.Text = "Select Branch"
    End Sub

    Private Sub lstboxTitle_SelectedIndexChanged(sender As Object, e As EventArgs) Handles lstboxTitle.SelectedIndexChanged
        If (lstboxTitle.SelectedIndex.Equals(0) Or lstboxTitle.SelectedIndex.Equals(3)) Then
            lstboxSex.SelectedIndex = 0
        End If
        If (lstboxTitle.SelectedIndex.Equals(1) Or lstboxTitle.SelectedIndex.Equals(2) Or lstboxTitle.SelectedIndex.Equals(4)) Then
            lstboxSex.SelectedIndex = 1
        End If
    End Sub
End Class
