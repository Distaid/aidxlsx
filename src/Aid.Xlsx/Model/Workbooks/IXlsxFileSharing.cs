// Keep this file CodeMaid organised and cleaned
namespace Aid.Xlsx.Model.Workbooks;

public interface IXlsxFileSharing
{
    //String AlgorithmName { get; set; }
    //Byte[] HashValue { get; set; }
    Boolean ReadOnlyRecommended { get; set; }

    //Byte[] ReservationPassword { get; set; }
    //Byte[] SaltValue { get; set; }
    //Int32 SpinCount { get; set; }
    String? UserName { get; set; }
}
