using System;
using System.Windows.Forms;
using System.Drawing;

namespace compare_exe
{
    internal static class MsgBox
    {
        public static DialogResult ShowWarn(string pMessage)
        {
            return MessageBox.Show(pMessage, "Warn", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Exclamation);
        }

        public static DialogResult ShowInfo(string pMessage)
        {
            return MessageBox.Show(pMessage, "Info", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Information);
        }

        public static DialogResult ShowError(string pMessage)
        {
            return MessageBox.Show(pMessage, "Error", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);
        }

        public static DialogResult Show(string promptText, string pMessage)
        {
            Form form = new Form();
            Label label = new Label();
            TextBox textBox = new TextBox();
            Button buttonOk = new Button();

            form.Text = "Info";
            label.Text = promptText;
            textBox.Text = pMessage;
            textBox.ReadOnly = true;

            buttonOk.Text = "OK";
            buttonOk.DialogResult = DialogResult.OK;

            label.SetBounds(3, 4, 35, 13);
            textBox.SetBounds(4, 20, 503, 20);
            buttonOk.SetBounds(209, 46, 93, 25);

            label.AutoSize = true;
            textBox.Anchor = textBox.Anchor | AnchorStyles.Right;
            buttonOk.Anchor = AnchorStyles.Bottom | AnchorStyles.Right;

            form.ClientSize = new Size(511, 78);
            form.Controls.AddRange(new Control[] { label, textBox, buttonOk });
            form.FormBorderStyle = FormBorderStyle.FixedDialog;
            form.StartPosition = FormStartPosition.CenterScreen;
            form.MinimizeBox = false;
            form.MaximizeBox = false;
            form.AcceptButton = buttonOk;

            DialogResult dialogResult = form.ShowDialog();
            return dialogResult;
        }


    }
}
