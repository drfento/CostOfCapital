using System;

namespace CostOfCapital
{
    class GetUserInfo
    {
        //Fields
        public string SignOnName;
        public DateTime SignOnTime;

        //Methods
        public void SetUserInfo(string signOnName, DateTime signOnTime)
        {
            SignOnName = signOnName;
            SignOnTime = signOnTime;
        }

    }
}
