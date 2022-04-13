using JumboExcel.Structure;
using NUnit.Framework;

namespace JumboExcel;

public class RelaxedCompatibilityTests
{
    [TestCase(31, false, true)]
    [TestCase(31, true, true)]
    [TestCase(32, false, false)]
    [TestCase(32, true, true)]
    public void CanUseVariousWorksheetNameLength(int length, bool allowLongWorksheetLength, bool successExpected)
    {
        try
        {
            _ = new Worksheet(new string('x', length),
                new WorksheetParametersElement(
                    false,
                    false,
                    null,
                    null,
                    allowLongWorksheetLength
                        ? WorksheetCompatibilityFlags.RELAX_WORKSHEET_LENGTH_CONSTRAINT
                        : WorksheetCompatibilityFlags.NONE));
            if (successExpected)
            {
                return;
            }
        }
        catch when (!successExpected)
        {
            return;
        }
        Assert.Fail($"Exception expected. Length {length} should not be acceptable in this case.");
    }
    
    [TestCase(31, false, true)]
    [TestCase(31, true, true)]
    [TestCase(32, false, false)]
    [TestCase(32, true, true)]
    public void CanUseVariousProgressingWorksheetNameLength(int length, bool allowLongWorksheetLength, bool successExpected)
    {
        try
        {
            _ = new ProgressingWorksheet<bool>(new string('x', length),
                new WorksheetParametersElement(
                    false,
                    false,
                    null,
                    null,
                    allowLongWorksheetLength
                        ? WorksheetCompatibilityFlags.RELAX_WORKSHEET_LENGTH_CONSTRAINT
                        : WorksheetCompatibilityFlags.NONE), _ => null);
            if (successExpected)
            {
                return;
            }
        }
        catch when (!successExpected)
        {
            return;
        }
        Assert.Fail($"Exception expected. Length {length} should not be acceptable in this case.");
    }
}