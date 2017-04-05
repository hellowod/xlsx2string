using System.Data;

/***
 * IChecker.cs
 * 
 * Author abaojin
 * Version 1.0
 * Date 2017.04.05
 */
namespace xlsx2string
{
    public interface IChecker
    {
        void Check(DataTable table);
    }
}
