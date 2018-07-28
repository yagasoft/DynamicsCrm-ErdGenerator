using System;
using CRM_ERD_Generator_GUI.UI;

namespace CRM_ERD_Generator_GUI.Helpers
{
    public static class Status
    {
	    public static Login Login;

        public static void Update(string message)
        {
			Login.Dispatcher.BeginInvoke(new Action(() => { Login.TextBoxLog.Text += message + "\n"; }));
        }

        public static void Clear()
        {
			Login.Dispatcher.BeginInvoke(new Action(() => { Login.TextBoxLog.Text = ""; }));
		}
    }
}
