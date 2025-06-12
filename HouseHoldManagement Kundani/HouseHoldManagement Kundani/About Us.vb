Public Class About_Us
    Private Sub About_Us_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ' Start the timer when form loads
        Timer1.Start()
        TextBox1.Text = "HouseholdDocument"
        TextBox2.Text = " Manager v1.0"
        ' Description of the app
        RichTextBox1.Text = "This application helps users organize, categorize, and manage household documents such as PDFs, images, Word, and Excel files. It includes features like previewing, filtering by category, and household linking."
        ' Contributors or team roles
        TextBox3.Text = "Organize and manage household documents efficiently."
        ListBox1.Items.Add("John Doe - Lead Developer")
        ListBox1.Items.Add("Jane Smith - UI/UX Designer")
        ListBox1.Items.Add("Alex Brown - QA Engineer")
        ' Load an image for the PictureBox (replace with actual image path if needed)
        ' Example: pbLogo.Image = Image.FromFile("C:\Path\To\Logo.png")
        ' You can also set the image in the designer.
        ' Link to contact or GitHub
        LinkLabel1.Text = "Visit our GitHub"
        LinkLabel1.Links.Add(0, LinkLabel1.Text.Length, "https://github.com/yourrepo")
    End Sub
    Private Sub LinkLabel1_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles LinkLabel1.LinkClicked
        'Process.Start(e.Link.LinkData.ToString())
        Try
            Process.Start(New ProcessStartInfo(e.Link.LinkData.ToString()) With {.UseShellExecute = True})
        Catch ex As Exception
            MessageBox.Show("Unable to open the link.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Me.Close()
    End Sub

    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        ' Stop the timer after first tick
        Timer1.Stop()
    End Sub
End Class