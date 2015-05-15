Imports System.Web
Imports CMS_API.WebUI.WebControls
Imports System.Data.SqlClient
Imports System.Web.UI.WebControls
Imports System.Web.UI.ClientScriptManager
Imports System.Net.Mail

Namespace SPSSDownloadForm

    Public Class SPSSDownloadForm

        Inherits BasePlaceHolderControl

        Dim portaluser
        Dim formPanel As New Panel
        Dim completedPanel As New Panel

        Dim WithEvents acceptCheckBox As New CheckBox

        Dim WithEvents submitButton As Button
        Dim WithEvents customValidatorAccept As CustomValidator

        Private request As HttpRequest
https://github.com/HelenSharmaSolent/SPSSDownloadForm/edit/master/SPSSDownloadForm.vb#        Private response As HttpResponse

        Public Sub addToOutput(ByVal controlIn As Object)
            Me.Controls.Add(controlIn)
        End Sub
        
        public function dothis()
        dim t as string = ""
        return t
        end function

        Private Sub SPSSDownloadForm_Init(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Init

            request = HttpContext.Current.Request
            response = HttpContext.Current.Response

            portaluser = DirectCast(Me.Page.FindControl("PortalHelper"), ljm.portal2.Security.PortalHelper).getPortalUser()

            formPanel.ID = "formpanel"
            completedPanel.ID = "completedpanel"

            addToOutput(formPanel)
            addToOutput(completedPanel)

            Me.FindControl("formpanel").Visible = True
            Me.FindControl("completedpanel").Visible = False

            Dim formFieldset As New FieldSet
            formFieldset.ID = "spssdownloadformfieldset"

            If PortalUser.getGroup = "public" Then
                formPanel.Controls.Add(New Span("You must be logged in as a student or staff member of access this form"))

            Else
                Try
                    Dim nameTextBox As New TextBox()
                    nameTextBox.Text = PortalUser.getFirstName + " " + PortalUser.getSurname
                    nameTextBox.ID = "nametextbox"
                    nameTextBox.ValidationGroup = "spssdownloadformgroup"
                    nameTextBox.MaxLength = "100"

                    Dim nameLabel As New Label
                    nameLabel.Text = "Requester Name:"
                    nameLabel.AssociatedControlID = nameTextBox.ID
                    nameLabel.Width = "200"

                    Dim nameDiv As New Div()
                    nameDiv.Controls.Add(nameLabel)
                    nameDiv.Controls.Add(nameTextBox)
                    nameDiv.Attributes.Add("style", "clear:both")

                    Dim emailTextBox As New TextBox()
                    emailTextBox.Text = PortalUser.getEmail
                    emailTextBox.ID = "emailtextbox"
                    emailTextBox.ValidationGroup = "spssdownloadformgroup"
                    emailTextBox.MaxLength = "50"

                    Dim emailLabel As New Label
                    emailLabel.Text = "University email address:"
                    emailLabel.AssociatedControlID = emailTextBox.ID
                    emailLabel.Width = "200"

                    Dim emailDiv As New Div()
                    emailDiv.Controls.Add(emailLabel)
                    emailDiv.Controls.Add(emailTextBox)
                    emailDiv.Attributes.Add("style", "clear:both")


                    Dim courseTextBox As New TextBox()
                    courseTextBox.Text = ""
                    courseTextBox.ID = "coursetextbox"
                    courseTextBox.ValidationGroup = "spssdownloadformgroup"
                    courseTextBox.MaxLength = "100"

                    Dim courseLabel As New Label
                    courseLabel.Text = "Course:"
                    courseLabel.AssociatedControlID = courseTextBox.ID
                    courseLabel.Width = "200"

                    Dim courseDiv As New Div()
                    courseDiv.Controls.Add(courseLabel)
                    courseDiv.Controls.Add(courseTextBox)
                    courseDiv.Attributes.Add("style", "clear:both")

                    Dim facultyTextBox As New TextBox()
                    facultyTextBox.Text = ""
                    facultyTextBox.ID = "facultytextbox"
                    facultyTextBox.ValidationGroup = "spssdownloadformgroup"
                    facultyTextBox.MaxLength = "100"

                    Dim facultyLabel As New Label
                    facultyLabel.Text = "Faculty/service:"
                    facultyLabel.AssociatedControlID = facultyTextBox.ID
                    facultyLabel.Width = "200"

                    Dim facultyDiv As New Div()
                    facultyDiv.Controls.Add(facultyLabel)
                    facultyDiv.Controls.Add(facultyTextBox)
                    facultyDiv.Attributes.Add("style", "clear:both")


                    Dim yearTextBox As New TextBox()
                    yearTextBox.Text = ""
                    yearTextBox.ID = "yeartextbox"
                    yearTextBox.ValidationGroup = "spssdownloadformgroup"
                    yearTextBox.Width = "50"
                    yearTextBox.MaxLength = "4"

                    Dim yearLabel As New Label
                    yearLabel.Text = "Year of study:"
                    yearLabel.AssociatedControlID = yearTextBox.ID
                    yearLabel.Width = "200"

                    Dim yearDiv As New Div()
                    yearDiv.Controls.Add(yearLabel)
                    yearDiv.Controls.Add(yearTextBox)
                    yearDiv.Attributes.Add("style", "clear:both")

                    acceptCheckBox.ID = "acceptcheckbox"
                    acceptCheckBox.ValidationGroup = "spssdownloadformgroup"
                    acceptCheckBox.Checked = False

                    Dim acceptLabel As New Label
                    acceptLabel.Text = "I accept the terms and conditions:"
                    acceptLabel.AssociatedControlID = acceptCheckBox.ID

                    Dim acceptDiv As New Div()
                    acceptDiv.Controls.Add(acceptLabel)
                    acceptDiv.Controls.Add(acceptCheckBox)
                    acceptDiv.Attributes.Add("style", "clear:both")


                    Dim requiredFieldName As New RequiredFieldValidator
                    requiredFieldName.ID = "rfvName"
                    requiredFieldName.ControlToValidate = nameTextBox.ID
                    requiredFieldName.Display = ValidatorDisplay.Dynamic
                    requiredFieldName.ValidationGroup = "spssdownloadformgroup"
                    requiredFieldName.ErrorMessage = "<span style=""color:red"">Please enter your name</span>"

                    Dim requiredFieldEmail As New RequiredFieldValidator
                    requiredFieldEmail.ID = "rfvEmail"
                    requiredFieldEmail.ControlToValidate = emailTextBox.ID
                    requiredFieldEmail.Display = ValidatorDisplay.Dynamic
                    requiredFieldEmail.ValidationGroup = "spssdownloadformgroup"
                    requiredFieldEmail.ErrorMessage = "<span style=""color:red"">Please enter your university email</span>"

                    Dim requiredFieldcourse As New RequiredFieldValidator
                    requiredFieldcourse.ID = "rfvcourse"
                    requiredFieldcourse.ControlToValidate = courseTextBox.ID
                    requiredFieldcourse.Display = ValidatorDisplay.Dynamic
                    requiredFieldcourse.ValidationGroup = "spssdownloadformgroup"
                    requiredFieldcourse.ErrorMessage = "<span style=""color:red"">Please enter your course</span>"

                    Dim requiredFieldyear As New RequiredFieldValidator
                    requiredFieldyear.ID = "rfvyear"
                    requiredFieldyear.ControlToValidate = yearTextBox.ID
                    requiredFieldyear.Display = ValidatorDisplay.Dynamic
                    requiredFieldyear.ValidationGroup = "spssdownloadformgroup"
                    requiredFieldyear.ErrorMessage = "<span style=""color:red"">Please enter your year of study</span>"

                    Dim requiredFieldfaculty As New RequiredFieldValidator
                    requiredFieldfaculty.ID = "rfvfaculty"
                    requiredFieldfaculty.ControlToValidate = facultyTextBox.ID
                    requiredFieldfaculty.Display = ValidatorDisplay.Dynamic
                    requiredFieldfaculty.ValidationGroup = "spssdownloadformgroup"
                    requiredFieldfaculty.ErrorMessage = "<span style=""color:red"">Please enter your faculty/service</span>"

                    Dim customValidatorAccept As New CustomValidator
                    customValidatorAccept.ID = "customValidatorAccept"
                    'customValidatorAccept.ControlToValidate = "acceptcheckbox"
                    customValidatorAccept.Display = ValidatorDisplay.Dynamic
                    customValidatorAccept.ValidationGroup = "spssdownloadformgroup"
                    customValidatorAccept.ErrorMessage = "<span style=""color:red"">Please confirm you have read the terms and conditions</span>"


                    submitButton = New Button
                    submitButton.ID = "updatebutton"
                    submitButton.Text = "Submit"
                    submitButton.CausesValidation = True
                    submitButton.ValidationGroup = "spssdownloadformgroup"

                    AddHandler submitButton.Click, AddressOf submitForm_Click
                    AddHandler customValidatorAccept.ServerValidate, AddressOf CustomValidator1_ServerValidate

                    Dim updateButtonDiv As New Div
                    updateButtonDiv.CssClass = "sys_form-buttons"
                    updateButtonDiv.Controls.Add(submitButton)

                    formFieldset.Controls.Add(New P("Please complete all fields"))
                    formFieldset.Controls.Add(nameDiv)
                    formFieldset.Controls.Add(requiredFieldName)
                    formFieldset.Controls.Add(emailDiv)
                    formFieldset.Controls.Add(requiredFieldEmail)

                    If portaluser.getGroup = "students" Then
                        formFieldset.Controls.Add(courseDiv)
                        formFieldset.Controls.Add(requiredFieldcourse)
                        formFieldset.Controls.Add(yearDiv)
                        formFieldset.Controls.Add(requiredFieldyear)
                    ElseIf portaluser.getGroup = "staff" Then
                        formFieldset.Controls.Add(facultyDiv)
                        formFieldset.Controls.Add(requiredFieldfaculty)
                    End If


                    formFieldset.Controls.Add(New P("&nbsp;"))
                    formFieldset.Controls.Add(New P("Please tick the checkbox below to confirm that you have read the <a href=""http://portal.solent.ac.uk/it-and-media/student-it-help/software-downloads/spss-software/resources/spss-academic-licence-terms-conditions.pdf"">terms and conditions</a>"))
                    formFieldset.Controls.Add(acceptDiv)
                    formFieldset.Controls.Add(customValidatorAccept)
                    formFieldset.Controls.Add(updateButtonDiv)

                    formPanel.Controls.Add(formFieldset)


                Catch ex As Exception
                    HttpContext.Current.Response.Write("ERROR - SPSSDownloadForm_Init: " & ex.Message())

                End Try
            End If
        End Sub


        Protected Sub submitForm_Click(ByVal sender As Object, ByVal e As System.EventArgs)
            Page.Validate("spssdownloadformgroup")
            'If acceptCheckBox.Checked = False Then
            '    e. = False
            'End If
            If Page.IsValid() Then
          
                Try

                    Dim nameTextBox As TextBox = FindControl("nametextbox")
                    Dim nameText As String = nameTextBox.Text

                    Dim emailTextBox As TextBox = FindControl("emailtextbox")
                    Dim emailText As String = emailTextBox.Text

                    Dim courseText As String = ""
                    Dim yearText As String = ""
                    Dim facultyText As String = ""

                    If portaluser.getGroup = "students" Then

                        Dim courseTextBox As TextBox = FindControl("coursetextbox")
                        courseText = courseTextBox.Text

                        Dim yearTextBox As TextBox = FindControl("yeartextbox")
                        yearText = yearTextBox.Text

                    ElseIf portaluser.getGroup = "staff" Then

                        Dim facultyTextBox As TextBox = FindControl("facultytextbox")
                        facultyText = facultyTextBox.Text
                    End If


                    Dim myDatabaseUtil As New edc.ljm.common.DatabaseUtil(edc.ljm.common.SiteUtil.DB_PORTAL_LIVE)

                    myDatabaseUtil.addParameter("@name", nameText)
                    myDatabaseUtil.addParameter("@email", emailText)

                    If portaluser.getGroup = "students" Then
                        myDatabaseUtil.addParameter("@course", courseText)
                        myDatabaseUtil.addParameter("@year", yearText)
                    ElseIf portaluser.getGroup = "staff" Then
                        myDatabaseUtil.addParameter("@course", facultyText)
                        myDatabaseUtil.addParameter("@year", "")
                    End If

                    Dim myDataReader As SqlDataReader = myDatabaseUtil.getSqlDataReader("storProc_StudentITSPSSRequestsInsert")

                    ljm.common.DatabaseUtil.closeSqlDataReader(myDataReader)

                    Dim completedDiv As New Div("<p>Thank you for your request.</p><p><strong>Version 22 Site Authorisation Code: 590f94742523616ccd2d</strong></p><p>Please download the sofware using the relevant link below.</p><ul><li><a href=""http://portal.solent.ac.uk/it-and-media/student-it-help/computer-discounts/spss/spss-downloads/spss-v22-32bit.exe"">SPSS Windows 32bit</a></li><li><a href=""http://portal.solent.ac.uk/it-and-media/student-it-help/computer-discounts/spss/spss-downloads/spss-v22-64bit.exe"">SPSS Windows 64bit</a></li><li><a href=""http://portal.solent.ac.uk/it-and-media/student-it-help/computer-discounts/spss/spss-downloads/spss-v22-mac.zip"">SPSS Mac</a></li></ul>")
                    completedPanel.Controls.Add(completedDiv)

                    Me.FindControl("formpanel").Visible = False
                    Me.FindControl("completedpanel").Visible = True

                Catch ex As Exception
                    response.Write("ERROR - SPSSDownloadForm submitForm_Click: " & ex.Message())

                End Try
            End If

        End Sub

        Protected Sub CustomValidator1_ServerValidate(ByVal source As Object, ByVal args As System.Web.UI.WebControls.ServerValidateEventArgs) 'Handles customValidatorAccept.ServerValidate
            If Not acceptCheckBox.Checked Then
                args.IsValid = False
            End If
        End Sub

    End Class

End Namespace
