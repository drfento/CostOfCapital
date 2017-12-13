using System;
using System.Globalization;
using static CostOfCapital.GetDisplay;

namespace CostOfCapital
{
    class GetSQL
    {
        //Fields
        public string SSQL;

        //Cost of capital view
        public string Cost_of_capital_View(CapitalFrm cform)
        {
            // Text Formated as 'Month[0] Year[1]'
            String[] dateText = cform.dateCB.Text.Split(' ');

            SSQL = " SELECT C.trader_description Trader, P.Portfolio,   "
             + "SUM(C.mtd_credit_charges) MTD, SUM(C.ytd_credit_charges) YTD, P.FUND_ID  FUND  "
             + "FROM table(RPTTRADE.populate_cost_of_cap('" + DateTime.ParseExact(dateText[0], "MMMM", CultureInfo.CurrentCulture).Month + "','" + dateText[1] + "')) C  "
             + " INNER Join (Select distinct t.TRADER_ID, t.PORTFOLIO_LABEL Portfolio, trim(UPPER(t.fund_ID)) FUND_ID, F.NAME   "
             + "        FROM ODSTRADE.TRADER_SUB_OBJECTIVE t   "
             + "        Inner join ODSTRADE.fund f on F.FUND_ID= T.FUND_ID   "
             + "        Where t.ACTIVE_FLAG = 'Y' and  T.FUND_ID in (Select FUND_ID From OPERATIONS.VW_FUND_ACTIVE)   "
             + "              AND t.PORTFOLIO_LABEL is not null   "
             + "    ) P on P.TRADER_ID=C.TRADER_ID and P.Name=C.fund_NAME   "
             + " Where P.FUND_ID not in ('XH')   "
             + "GROUP BY C.trader_description,P.Portfolio, P.FUND_ID    "
             + "order by  C.trader_description   ";

            return SSQL;

        }

        //Lookup sent date
        public string Lookup_Sent_Info(string unid)
        {
            SSQL = " SELECT TYPOLOGY, PORTFOLIO, CPTY, INST, FX, TPR,  TDATE, TAMOUNT, TFX, '' TTRANS, TCOMMENT, STRATEGY, "
             + " '' BRKTYPE, BRKSRCE, BRKDEST, UNIQUE_ID, FUND, ACCT_NUM, Sent_ON, SENT_AMT  "
             + " FROM OPERATIONS.COST_OF_CAPITAL_SCF  "
             + " Where SubStr(Portfolio || ':' || TCOMMENT,1,31) = '" + unid + "'  ";

            return SSQL;

        }

        //Cost of Capital Merge
        public string CapitalCost_of_capital_Merge(CapitalFrm cform, int i)
        {
            //(string)cform.dgvCostView.Rows[i].Cells[(int)Capital.portfolio].Value
            SSQL = "Merge INTO OPERATIONS.COST_OF_CAPITAL_SCF t "
            + " Using(  "
            + " Select '" + (string)cform.dgvFeedView.Rows[i].Cells[(int)Capital.fTcomment].Value + "' TCOMMENT, " 
            + "        '" + (string)cform.dgvFeedView.Rows[i].Cells[(int)Capital.fPortfolio].Value + "' PORTFOLIO "
            + " From DUAL ) d "
            + " on (d.TCOMMENT = t.TCOMMENT and d.PORTFOLIO = t.PORTFOLIO) "
            + " When NOT MATCHED THEN "
            + " Insert  ( "
            + "   t.TYPOLOGY, t.PORTFOLIO, t.CPTY, t.INST, t.FX, t.TPR, "
            + "   t.TDATE, t.TAMOUNT, t.TFX, t.TCOMMENT, t.STRATEGY, t.BRKSRCE,  "
            + "   t.BRKDEST, t.UNIQUE_ID, t.FUND, t.ACCT_NUM, t.SENT_AMT)  "
            + "VALUES ( "
            + " '" + (string)cform.dgvFeedView.Rows[i].Cells[(int)Capital.fTypology].Value + "', " // Typology
            + " '" + (string)cform.dgvFeedView.Rows[i].Cells[(int)Capital.fPortfolio].Value + "', " // Portfolio
            + " '" + (string)cform.dgvFeedView.Rows[i].Cells[(int)Capital.fCpty].Value + "', " // CPTY
            + " '" + (string)cform.dgvFeedView.Rows[i].Cells[(int)Capital.fInst].Value + "', " // INST
            + " '" + (string)cform.dgvFeedView.Rows[i].Cells[(int)Capital.fFx].Value + "', " // FX
            + " '" + (string)cform.dgvFeedView.Rows[i].Cells[(int)Capital.fTpr].Value + "', " // TPR
            + " TO_DATE('" + (string)cform.dgvFeedView.Rows[i].Cells[(int)Capital.fTdate].Value + "', 'MM/DD/YYYY'), " // TDATE
            + " " + Convert.ToDouble(cform.dgvFeedView.Rows[i].Cells[(int)Capital.fTamount].Value) + ", " // TAMOUNT
            + " '" + (string)cform.dgvFeedView.Rows[i].Cells[(int)Capital.fTfx].Value + "', " // TFX
            + " '" + (string)cform.dgvFeedView.Rows[i].Cells[(int)Capital.fTcomment].Value + "', " // TCOMMENT
            + " '" + (string)cform.dgvFeedView.Rows[i].Cells[(int)Capital.fStrategy].Value + "', " // STRATEGY
            + " '" + (string)cform.dgvFeedView.Rows[i].Cells[(int)Capital.fBrksrce].Value + "', " // BRKSRCE
            + " '" + (string)cform.dgvFeedView.Rows[i].Cells[(int)Capital.fBrkdest].Value + "', " // BRKDEST
            + " '" + (string)cform.dgvFeedView.Rows[i].Cells[(int)Capital.fUnique_Id].Value + "', " // UNIQUE_ID
            + " '" + (string)cform.dgvFeedView.Rows[i].Cells[(int)Capital.fFund].Value + "', " // FUND
            + " '" + (string)cform.dgvFeedView.Rows[i].Cells[(int)Capital.fAcct_Num].Value + "', " // ACCT_NUM
            + " " + Convert.ToDouble(cform.dgvFeedView.Rows[i].Cells[(int)Capital.fSent_Amt].Value) + " " // SENT_AMT
            + " ) ";

            return SSQL;

        }

        //Update sent on if it has not already been sent
        public string Final_Sent_Merge(CapitalFrm cform, int i)
        {

            string tprString = "";

            if ((string)cform.dgvFinalView.Rows[i].Cells[(int)Capital.fTpr].Value == "PAY")
            {
                tprString = "  t.Sent_AMT = t.Sent_AMT - " + Convert.ToDouble(cform.dgvFeedView.Rows[i].Cells[(int)Capital.fTamount].Value);
            }
            else
            {
                tprString = " t.Sent_AMT = t.Sent_AMT + " + Convert.ToDouble(cform.dgvFeedView.Rows[i].Cells[(int)Capital.fTamount].Value);
            }


            SSQL = "Merge INTO OPERATIONS.COST_OF_CAPITAL_SCF t "
            +" Using(  "
            + " Select '" + (string)cform.dgvFinalView.Rows[i].Cells[(int)Capital.fTcomment].Value + "' TCOMMENT, "
            + " '       " + (string)cform.dgvFinalView.Rows[i].Cells[(int)Capital.fPortfolio].Value + "' PORTFOLIO "
            + " From DUAL ) d "
            + " on (d.TCOMMENT = t.TCOMMENT and d.PORTFOLIO = t.PORTFOLIO) "
            + " When MATCHED THEN "
            + " UPDATE SET  t.Sent_ON = TO_DATE(SYSDATE),  " + tprString;

            return SSQL;

        }

    }
}
