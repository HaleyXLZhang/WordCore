using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace WordCoreTests.LetterGenerationTest
{
    public class ReasonInfo
    {
        public string FileName { get; set; }
        public List<SelectReasonItem> Reasons;
        public ReasonInfo()
        {
            Reasons = new List<SelectReasonItem>();

        }
    }
   public class SelectReasonItem
    {
        public int Row { get; set; }
        public int CopyColumn { get; set; }
        public string Code { get; set; }
        public string ShortCode { get; set; }
    }
}
