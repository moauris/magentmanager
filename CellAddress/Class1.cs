using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace ControlMAgent.Base
{
    
    //Instantiate a CellAddress
    public class CellAddress
    {
        string fullAddress;
        char column;
        int currentRow, OriginalRow;
        public bool IsVIP { get => OriginalRow == 49; }
        //Summary
        //      Instantiate with current Cell Address string
        //      Sample: $A$11
        public CellAddress(string AddressString)
        {
            Regex rxValid = new Regex(@"^\$[A-Z]\$\d+$");
            if (!rxValid.IsMatch(AddressString)) //If not match, throw exception
                throw new Exception(AddressString
                    + " is not a valid Cell Address.");

            fullAddress = AddressString;
            Regex rxExtract = new Regex(
                @"^\$(? 'CStart'[A - Z])\$(? 'RStart'\d +)");
            column = rxExtract.Matches(AddressString)[0].Groups["CStart"].Value.ToCharArray()[0];
            currentRow = OriginalRow = 
                int.Parse(rxExtract.Matches(AddressString)[0].Groups["RStart"].Value);

        }
        //Summary
        //      Instantiate with current Cell Address string
        //      Sample: Col = A, Row = 11
        public CellAddress(char Col, int Row)
        {
            string AddressString = string
                .Format("${0}${1}", Col, Row);
            new CellAddress(AddressString);
        }

        public CellAddress(string StartCell, string EndCell)
        {
            string Address = StartCell + ":" + EndCell;
            fullAddress = Address;
        }
        // Summary
        //      Get the Full Address String
        public string FullAddress()
        {
            return fullAddress;
        }
        // Summary
        //      Increment one for Row
        public string NextRow
        {
            get
            {
                currentRow += 1;
                fullAddress = string.Format("${0}${1}", column, currentRow);
                return fullAddress;
            }
        }
        public char GetCurrentCol{ get => column; }
        public int GetCurrentRow { get => currentRow; }
        
    }

    public class CellAddresses
    {
        string fullAddress;
        char column;
        int CurrentRow, OriginalRow;
        public bool IsVIP { get => OriginalRow == 49; }
        string fullAddressList;
        //Summary
        //      Instantiate with current Cell Address string
        //      Sample: $A$11:$B$13
        public CellAddresses(string AddressString)
        {
            Match mValid = Regex.Match(AddressString
                , @"^\$[A-Z]\$\d+:\$[A-Z]\$\d+$");
            if (!mValid.Success)
                throw new Exception(AddressString 
                    + " is not a valid Cell Address Area.");
            Regex rxExtract = new Regex(
                @"^\$(?'CStart'[A-Z])\$(?'RStart'\d+):\$(?'CEnd'[A-Z])\$(?'REnd'\d+)$");
            char CStart, CEnd;
            int RStart, REnd;
            CStart = rxExtract.Matches(AddressString)[0]
                .Groups["CStart"].Value.ToCharArray()[0];
            RStart = int.Parse(rxExtract.Matches(AddressString)[0]
                .Groups["RStart"].Value);
            CEnd = rxExtract.Matches(AddressString)[0].Groups["CEnd"]
                .Value.ToCharArray()[0];
            REnd = int.Parse(rxExtract.Matches(AddressString)[0]
                .Groups["REnd"].Value);
            //Debug.Print("CStart={0};CEnd={1};RStart={2};REnd={3}"
            string AddressRange = "";
            for (char c = CStart; c <= CEnd; c++)
            {
                for (int i = RStart; i <= REnd; i++)
                {
                    string thisRange = string.Format("${0}${1}", c, i);
                    AddressRange += thisRange + ";";
                }
            }
            fullAddressList = AddressRange;

        }
        public CellAddresses(CellAddress CellStart, CellAddress CellEnd)
        {
            string AddressString = string.Format("{0}:{1}"
                , CellStart.FullAddress()
                , CellEnd.FullAddress());
            new CellAddresses(AddressString);
        }
            //Summary
            //      Instantiate with current Cell Address string
            //      Sample: Start Cell = $A$11 Extend Row = 3
            public CellAddresses(string StartCell, int ExtendRow)
        {

            char CStart, CEnd;
            int RStart, REnd;
            CellAddress CellStart = new CellAddress(StartCell);
            CStart = CellStart.GetCurrentCol;
            RStart = CellStart.GetCurrentRow;
            CEnd = CStart;
            for (int i = 1; i < 4; i++) CEnd++;
            REnd = RStart + ExtendRow;
            CellAddress CellEnd = new CellAddress(CEnd, REnd);
            new CellAddresses(CellStart, CellEnd);

        }

        // Summary
        //      Get the Full Address String, such as $A$8:$B$12
        public string FullAddress()
        {
            return fullAddress;
        }
        // Summary
        //      Get the iteration of the Cell Area;
        //      $H$25;$H$26;$H$27;...$L$74;$L$75;$L$76;
        public string FullAddressList()
        {
            return fullAddressList;
        }

    }

}
