using System;
using System.Data.Odbc;
using System.Globalization;
using System.Drawing;
using System.Linq;

namespace CostOfCapital
{
    class GetDisplay
    {
        public enum Capital
        {   //Capital
            trader = 0,
            portfolio, 
            mtd,
            ytd,
            fund,
            //Feed & Final
            fTypology = 0,
            fPortfolio,
            fCpty,
            fInst,
            fFx,
            fTpr,
            fTdate,
            fTamount,
            fTfx,
            fTtrans,
            fTcomment,
            fStrategy,
            fBrktype,
            fBrksrce,
            fBrkdest,
            fUnique_Id,
            fFund,
            fAcct_Num,
            fSent_On,
            fSent_Amt,
            fSent_Box
        }


        //Treasury Portfolios
        public string[] treasArray = { "CAXTON INTL", "QA-TREASURY", "QE-TREASURY", "QO-TREASURY", "QM-TREASURY" };

        public void AppInit(CapitalFrm cform)  //Main Setup for form
        {
            //Variables
            GetUserInfo user = new GetUserInfo();
            user.SetUserInfo(signOnName: Environment.UserName, signOnTime: DateTime.Now);

            //ComboBox
            DisplayComboBox(cform);

            //Log on info
            cform.userInfoTb.Text = user.SignOnName.ToUpper() + ": " + user.SignOnTime.ToShortDateString();
        }

        public void DisplayComboBox(CapitalFrm cform) //Populate the Month Year Combo Box
        {
            DateTime myDate = DateTime.Now.Date.AddMonths(-1);
            // Show last 6 months in combobox
            for (int i = 1; i < 7; i++)
            {
                cform.dateCB.Items.Add(myDate.ToString("MMMM yyyy", CultureInfo.InvariantCulture));
                myDate = myDate.AddMonths(-1);
            }
            // Show first index as default
            cform.dateCB.SelectedIndex = 0;
        }


        public void DisplayCapitalView(CapitalFrm cform) //Display DGV on Cost of Capital
        {
            //Clear all DGVs
            cform.dgvFinalView.Rows.Clear();
            cform.dgvFinalView.Refresh();
            cform.dgvFeedView.Rows.Clear();
            cform.dgvFeedView.Refresh();
            cform.dgvCostView.Rows.Clear();
            cform.dgvCostView.Refresh();
            // Variables
            Connection connection = new Connection();
            GetSQL sql = new GetSQL();
            OdbcDataReader dataReader = connection.RunQuery(sql.Cost_of_capital_View(cform));

            if (dataReader.HasRows)
            {
                while (dataReader.Read())
                {
                    cform.dgvCostView.Rows.Add(dataReader["TRADER"].ToString(), 
                                                 dataReader["PORTFOLIO"].ToString(), 
                                                 dataReader["MTD"],
                                                 dataReader["YTD"],
                                                 dataReader["FUND"].ToString()
                                                 );
                }

                //Add fund level totals
                double ytdCI, ytdQA, ytdQE, ytdQO, ytdQM;
                double mtdCI, mtdQA, mtdQE, mtdQO, mtdQM;
                ytdCI = ytdQA = ytdQE = ytdQO = ytdQM = 0;
                mtdCI = mtdQA = mtdQE = mtdQO = mtdQM = 0;

                for (int i = 0; i < cform.dgvCostView.Rows.Count; ++i)
                {
                    switch ((string)cform.dgvCostView.Rows[i].Cells[(int)Capital.fund].Value)
                    {
                        case "CI":
                            ytdCI += Convert.ToDouble(cform.dgvCostView.Rows[i].Cells[(int)Capital.ytd].Value);
                            mtdCI += Convert.ToDouble(cform.dgvCostView.Rows[i].Cells[(int)Capital.mtd].Value);
                            break;
                        case "QA":
                            ytdQA += Convert.ToDouble(cform.dgvCostView.Rows[i].Cells[(int)Capital.ytd].Value);
                            mtdQA += Convert.ToDouble(cform.dgvCostView.Rows[i].Cells[(int)Capital.mtd].Value);
                            break;
                        case "QE":
                            ytdQE += Convert.ToDouble(cform.dgvCostView.Rows[i].Cells[(int)Capital.ytd].Value);
                            mtdQE += Convert.ToDouble(cform.dgvCostView.Rows[i].Cells[(int)Capital.mtd].Value);
                            break;
                        case "QO":
                            ytdQO += Convert.ToDouble(cform.dgvCostView.Rows[i].Cells[(int)Capital.ytd].Value);
                            mtdQO += Convert.ToDouble(cform.dgvCostView.Rows[i].Cells[(int)Capital.mtd].Value);
                            break;
                        case "QM":
                            ytdQM += Convert.ToDouble(cform.dgvCostView.Rows[i].Cells[(int)Capital.ytd].Value);
                            mtdQM += Convert.ToDouble(cform.dgvCostView.Rows[i].Cells[(int)Capital.mtd].Value);
                            break;
                    } 
                }
                //Display totals if MTD not 0
                int nRowIndex = 0;

                if (mtdCI != 0 || ytdCI != 0)
                {
                    nRowIndex = cform.dgvCostView.Rows.Count;
                    cform.dgvCostView.Rows.Add("CI", "CAXTON INTL", mtdCI, ytdCI, "CI");
                    cform.dgvCostView.Rows[nRowIndex].DefaultCellStyle.BackColor = Color.Beige;
                }
                if (mtdQA != 0 || ytdQA != 0)
                {
                    nRowIndex = cform.dgvCostView.Rows.Count;
                    cform.dgvCostView.Rows.Add("QA", "QA-TREASURY", mtdQA, ytdQA, "QA");
                    cform.dgvCostView.Rows[nRowIndex].DefaultCellStyle.BackColor = Color.Beige;
                }
                if (mtdQE != 0 || ytdQE != 0)
                {
                    nRowIndex = cform.dgvCostView.Rows.Count;
                    cform.dgvCostView.Rows.Add("QE", "QE-TREASURY", mtdQE, ytdQE, "QE");
                    cform.dgvCostView.Rows[nRowIndex].DefaultCellStyle.BackColor = Color.Beige;
                }
                if (mtdQO != 0 || ytdQO != 0)
                {
                    nRowIndex = cform.dgvCostView.Rows.Count;
                    cform.dgvCostView.Rows.Add("QO", "QO-TREASURY", mtdQO, ytdQO, "QO");
                    cform.dgvCostView.Rows[nRowIndex].DefaultCellStyle.BackColor = Color.Beige;
                }
                if (mtdQM != 0 || ytdQM != 0)
                {
                    nRowIndex = cform.dgvCostView.Rows.Count;
                    cform.dgvCostView.Rows.Add("QM", "QM-TREASURY", mtdQM, ytdQM, "QM");
                    cform.dgvCostView.Rows[nRowIndex].DefaultCellStyle.BackColor = Color.Beige;
                }

            }
            // close
            dataReader.Close();
        }

        public void DisplayFeedView(CapitalFrm cform) //Display DGV on Feed tab
        {
            //Clear
            cform.dgvFeedView.Rows.Clear();
            cform.dgvFeedView.Refresh();

            Connection connection = new Connection();
            GetSQL sql = new GetSQL();
            //Variables
            string cpty, tpr, strat, sent_unid, sent_date;
            double amt, sent_amt;
            bool send_bool;


            //loop through capital view to create feed view
            for (int i = 0; i < cform.dgvCostView.Rows.Count; ++i)
            {
                //If MTD has a value add to feed
                if (Convert.ToDouble(cform.dgvCostView.Rows[i].Cells[(int)Capital.mtd].Value) != 0)
                {
                    //----Set variables
                    switch ((string)cform.dgvCostView.Rows[i].Cells[(int)Capital.trader].Value)
                    {
                        case "CI":
                            cpty = "CXI";
                            strat = "CAP USAGE";
                            break;
                        case "QA":
                        case "QE":
                        case "QO":
                        case "QM":
                            cpty = "SCF";
                            strat = "CAP USAGE";
                            break;
                        default:
                            cpty = "SCF";
                            strat = "ADJUSTMENTS";
                            break;
                    }

                    sent_unid = (string)cform.dgvCostView.Rows[i].Cells[(int)Capital.portfolio].Value + ":" + cform.dateCB.Text.ToUpper() + " CAP USAGE";
                    if (sent_unid.Length >= 31)
                    {
                        sent_unid = sent_unid.Substring(0, 31);
                    }
                    OdbcDataReader dataReader = connection.RunQuery(sql.Lookup_Sent_Info(sent_unid));

                    if (dataReader.HasRows)
                    {
                        sent_date = ((DateTime)dataReader["Sent_ON"]).ToString("MM/dd/yyyy");
                        sent_amt = Convert.ToDouble(dataReader["SENT_AMT"]);
                    }
                    else
                    {
                        sent_date = null;
                        sent_amt = 0;
                    }

                    if (Convert.ToDouble(cform.dgvCostView.Rows[i].Cells[(int)Capital.mtd].Value) > 0)
                    {
                        tpr = "RECEIVE";
                        amt = Convert.ToDouble(cform.dgvCostView.Rows[i].Cells[(int)Capital.mtd].Value);
                    }
                    else
                    {
                        tpr = "PAY";
                        amt = (Convert.ToDouble(cform.dgvCostView.Rows[i].Cells[(int)Capital.mtd].Value) * -1);
                    }

                    if (sent_date == null)
                    {
                        send_bool = true;
                    }
                    else if ((treasArray.Contains((string)cform.dgvCostView.Rows[i].Cells[(int)Capital.portfolio].Value) == false)
                            && (Math.Abs(amt) > Math.Abs(sent_amt)))
                    {
                        send_bool = true;
                    }
                    else
                    {
                        send_bool = false;
                    }

                    // Fill out feed DGV
                    cform.dgvFeedView.Rows.Add("IMPUTED INTEREST",  //Typology
                          (string)cform.dgvCostView.Rows[i].Cells[(int)Capital.portfolio].Value, //Portfolio
                          cpty, //cpty
                          "IMPIT", //Inst
                          "USD", //FX
                          tpr, //TPR
                          DateTime.Now.ToString("MM/dd/yyyy"), //TDATE
                          amt, //AMOUNT
                          "USD", //TFX
                          "", //TTRANS
                          cform.dateCB.Text.ToUpper() + " CAP USAGE", //TCOMMENT
                          strat, //STRATEGY
                          "", //BRKTYPE
                          (string)cform.dgvCostView.Rows[i].Cells[(int)Capital.portfolio].Value, //BRKSRC
                          "ADJLAA", //BRKDEST
                          "COSTOFCAP_SCF_" + DateTime.Now.ToString("yymmddhhMMss") + i.ToString(),  //UNID
                          (string)cform.dgvCostView.Rows[i].Cells[(int)Capital.fund].Value, //FUND
                          "CAXT", //ACCT
                          sent_date,
                          sent_amt,
                          send_bool
                         );
                    // close
                    dataReader.Close();
                }

            }

        }

        public void MergeFeedView(CapitalFrm cform) //Display DGV on Feed tab
        {
            Connection connection = new Connection();
            GetSQL sql = new GetSQL();
            //loop through feed view and use merge to insert if unmatched
            for (int i = 0; i < cform.dgvFeedView.Rows.Count; ++i)
            {

                connection.SendQuery(sql.CapitalCost_of_capital_Merge(cform,i)); 

            }
        }

        public void CreateFinalFeedView(CapitalFrm cform) //Display DGV on FINAL Feed tab
        {
            //Clear
            cform.dgvFinalView.Rows.Clear();
            cform.dgvFinalView.Refresh();

            //Variables
            string sent_unid;
            double sent_amt;
            bool fundCI, fundQA, fundQE, fundQO, fundQM;
            fundCI = fundQA = fundQE = fundQO = fundQM = false;

            Connection connection = new Connection();
            GetSQL sql = new GetSQL();

            //loop through feed view and use merge to insert if unmatched
            for (int i = 0; i < cform.dgvFeedView.Rows.Count; ++i)
            {
                //Construct Unid
                sent_unid = (string)cform.dgvFeedView.Rows[i].Cells[(int)Capital.portfolio].Value + ":" + cform.dateCB.Text.ToUpper() + " CAP USAGE";
                if (sent_unid.Length >= 31)
                {
                    sent_unid = sent_unid.Substring(0, 31);
                }
                OdbcDataReader dataReader = connection.RunQuery(sql.Lookup_Sent_Info(sent_unid));
                //Retrieve amount sent
                if (dataReader.HasRows)
                {
                    sent_amt = Convert.ToDouble(dataReader["SENT_AMT"]);
                }
                else
                {
                    sent_amt = 0;
                }

                //If box checked and it has not been sent an amount
                if ((treasArray.Contains((string)cform.dgvFeedView.Rows[i].Cells[(int)Capital.portfolio].Value) == false)
                    && ((bool)cform.dgvFeedView.Rows[i].Cells[(int)Capital.fSent_Box].Value == true)
                    && sent_amt == 0
                    )
                {

                    cform.dgvFinalView.Rows.Add(
                                    (string)cform.dgvFeedView.Rows[i].Cells[(int)Capital.fTypology].Value,  // Typology
                                    (string)cform.dgvFeedView.Rows[i].Cells[(int)Capital.fPortfolio].Value, // Portfolio
                                    (string)cform.dgvFeedView.Rows[i].Cells[(int)Capital.fCpty].Value, // CPTY
                                    (string)cform.dgvFeedView.Rows[i].Cells[(int)Capital.fInst].Value, // INST
                                    (string)cform.dgvFeedView.Rows[i].Cells[(int)Capital.fFx].Value, // FX
                                    (string)cform.dgvFeedView.Rows[i].Cells[(int)Capital.fTpr].Value, // TPR
                                    (string)cform.dgvFeedView.Rows[i].Cells[(int)Capital.fTdate].Value, // TDATE
                                    Convert.ToDouble(cform.dgvFeedView.Rows[i].Cells[(int)Capital.fTamount].Value), // TAMOUNT
                                    (string)cform.dgvFeedView.Rows[i].Cells[(int)Capital.fTfx].Value, // TFX
                                    (string)cform.dgvFeedView.Rows[i].Cells[(int)Capital.fTtrans].Value, // Ttrans
                                    (string)cform.dgvFeedView.Rows[i].Cells[(int)Capital.fTcomment].Value,  // TCOMMENT
                                    (string)cform.dgvFeedView.Rows[i].Cells[(int)Capital.fStrategy].Value, // STRATEGY
                                    (string)cform.dgvFeedView.Rows[i].Cells[(int)Capital.fBrktype].Value,  // BRKTYPE
                                    (string)cform.dgvFeedView.Rows[i].Cells[(int)Capital.fBrksrce].Value,  // BRKSRCE
                                    (string)cform.dgvFeedView.Rows[i].Cells[(int)Capital.fBrkdest].Value, // BRKDEST
                                    (string)cform.dgvFeedView.Rows[i].Cells[(int)Capital.fUnique_Id].Value, // UNIQUE_ID
                                    (string)cform.dgvFeedView.Rows[i].Cells[(int)Capital.fFund].Value, // FUND
                                    (string)cform.dgvFeedView.Rows[i].Cells[(int)Capital.fAcct_Num].Value // ACCT_NUM
                                              );

                    //set fund bools
                    switch ((string)cform.dgvFeedView.Rows[i].Cells[(int)Capital.fFund].Value)
                    {
                        case "CI":
                            fundCI = true;
                            break;
                        case "QA":
                            fundQA = true;
                            break;
                        case "QE":
                            fundQE = true;
                            break;
                        case "QO":
                            fundQO = true;
                            break;
                        case "QM":
                            fundQM = true;
                            break;
                    }

                }
                // close
                dataReader.Close();
            }

            //Call enrich feed for treasury 
            if (fundCI == true)
            {
                sent_unid = "CAXTON INTL" + ":" + cform.dateCB.Text.ToUpper() + " CAP USAGE";
                if (sent_unid.Length >= 31)
                {
                    sent_unid = sent_unid.Substring(0, 31);
                }
                EnrichFinalFeed(cform, sent_unid, "CI");
            }
            if (fundQA == true)
            {
                sent_unid = "QA-TREASURY" + ":" + cform.dateCB.Text.ToUpper() + " CAP USAGE";
                if (sent_unid.Length >= 31)
                {
                    sent_unid = sent_unid.Substring(0, 31);
                }
                EnrichFinalFeed(cform, sent_unid, "QA");
            }
            if (fundQE == true)
            {
                sent_unid = "QE-TREASURY" + ":" + cform.dateCB.Text.ToUpper() + " CAP USAGE";
                if (sent_unid.Length >= 31)
                {
                    sent_unid = sent_unid.Substring(0, 31);
                }
                EnrichFinalFeed(cform, sent_unid, "QE");
            }
            if (fundQO == true)
            {
                sent_unid = "QO-TREASURY" + ":" + cform.dateCB.Text.ToUpper() + " CAP USAGE";
                if (sent_unid.Length >= 31)
                {
                    sent_unid = sent_unid.Substring(0, 31);
                }
                EnrichFinalFeed(cform, sent_unid, "QO");
            }
            if (fundQM == true)
            {
                sent_unid = "QM-TREASURY" + ":" + cform.dateCB.Text.ToUpper() + " CAP USAGE";
                if (sent_unid.Length >= 31)
                {
                    sent_unid = sent_unid.Substring(0, 31);
                }
                EnrichFinalFeed(cform, sent_unid,"QM");
            }
            //Select tab
            cform.CapitalTabs.SelectTab(2);
        }

        public void EnrichFinalFeed(CapitalFrm cform, string sent_unid, string fund) //Add fund level totals with updated amounts
        {
            double tAmount;
            string tTpr;
            tAmount = 0;

            Connection connection = new Connection();
            GetSQL sql = new GetSQL();
            OdbcDataReader dataReader = connection.RunQuery(sql.Lookup_Sent_Info(sent_unid));

            //If fund matches update amount for items beign sent
            for (int i = 0; i < cform.dgvFinalView.Rows.Count; ++i)
            {
                if (fund == (string)cform.dgvFinalView.Rows[i].Cells[(int)Capital.fFund].Value)
                {
                  if ((string)cform.dgvFinalView.Rows[i].Cells[(int)Capital.fTpr].Value == "PAY")
                    {
                        tAmount = tAmount - Convert.ToDouble(cform.dgvFinalView.Rows[i].Cells[(int)Capital.fTamount].Value);
                    }
                     else
                    {
                        tAmount = tAmount + Convert.ToDouble(cform.dgvFinalView.Rows[i].Cells[(int)Capital.fTamount].Value);
                    }

                }
            }

            //Update direction of TPR
            if (tAmount > 0)
            {
                tTpr = "PAY";
            }
            else
            {
                tTpr = "RECEIVE";
            }

            int nRowIndex = 0;

            //loop through and update amounts
            //apply to feed
            if (dataReader.HasRows)
            {
                nRowIndex = cform.dgvFinalView.Rows.Count;
                cform.dgvFinalView.Rows.Add(
                           (string)dataReader["TYPOLOGY"],  // Typology
                           (string)dataReader["PORTFOLIO"], // Portfolio
                           (string)dataReader["CPTY"], // CPTY
                           (string)dataReader["INST"], // INST
                           (string)dataReader["FX"], // FX
                           tTpr, // TPR
                           ((DateTime)dataReader["TDATE"]).ToString("MM/dd/yyyy"), // TDATE
                           Math.Abs(tAmount), // TAMOUNT
                           (string)dataReader["TFX"], // TFX
                           dataReader["TTRANS"].ToString(), // Ttrans
                           (string)dataReader["TCOMMENT"],  // TCOMMENT
                           (string)dataReader["STRATEGY"], // STRATEGY
                           dataReader["BRKTYPE"].ToString(),  // BRKTYPE
                           (string)dataReader["BRKSRCE"],  // BRKSRCE
                           (string)dataReader["BRKDEST"], // BRKDEST
                           (string)dataReader["UNIQUE_ID"], // UNIQUE_ID
                           (string)dataReader["FUND"], // FUND
                           (string)dataReader["ACCT_NUM"]  // ACCT_NUM
                                     );
                cform.dgvFinalView.Rows[nRowIndex].DefaultCellStyle.BackColor = Color.Beige;
            }

            // close
            dataReader.Close();

        }
    }
}
