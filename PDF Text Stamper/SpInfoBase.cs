﻿namespace PDF_Text_Stamper
{
    public class SpInfoBase
    {
        public string SpFileName { get; set; }
        public string SpQuantity { get; set; }
        public string SpColor { get; set; }
        public string SpStatus { get; set; } = "Failed !";
        public string SpProject { get; set; }
        public string SpMepBy { get; set; }
    }
}