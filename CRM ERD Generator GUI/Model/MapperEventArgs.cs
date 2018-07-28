using System;

namespace CRM_ERD_Generator_GUI.Model
{
    public class MapperEventArgs : EventArgs
    {
        public string Message { get; set; }
        public string MessageExtended { get; set; }
    }
}
