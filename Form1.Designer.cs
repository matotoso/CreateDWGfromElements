partial class Form1
{
    private System.ComponentModel.IContainer components = null;

    protected override void Dispose(bool disposing)
    {
        if (disposing && (components != null))
        {
            components.Dispose();
        }
        base.Dispose(disposing);
    }

    private void InitializeComponent()
    {
        this.SuspendLayout();
        // 
        // Form1
        // 
        this.ClientSize = new System.Drawing.Size(800, 450);
        this.Name = "Form1";
        this.Text = "My Revit Add-in";
        this.Load += new System.EventHandler(this.Form1_Load);
        this.ResumeLayout(false);
    }
}
