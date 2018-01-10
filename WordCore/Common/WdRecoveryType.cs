using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace WordCore.Common
{
    public enum WdRecoveryType
    {
        wdPasteDefault = 0,
        wdSingleCellText = 5,
        wdSingleCellTable = 6,
        wdListContinueNumbering = 7,
        wdListRestartNumbering = 8,
        wdTableAppendTable = 10,
        wdTableInsertAsRows = 11,
        wdTableOriginalFormatting = 12,
        wdChartPicture = 13,
        wdChart = 14,
        wdChartLinked = 15,
        wdFormatOriginalFormatting = 16,
        wdUseDestinationStylesRecovery = 19,
        wdFormatSurroundingFormattingWithEmphasis = 20,
        wdFormatPlainText = 22,
        wdTableOverwriteCells = 23,
        wdListCombineWithExistingList = 24,
        wdListDontMerge = 25
    }
}
