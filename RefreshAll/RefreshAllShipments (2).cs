using Epicor.Mfg.BO;
using Epicor.Mfg.Core;
using Microsoft.Office.Interop.Outlook;
using OutlookApp = Microsoft.Office.Interop.Outlook.Application;
using System;
using System.Collections;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RefreshAll
{
    public class RefreshAllShipments
    {
        //main method
        static void Main()
        {
            //a tray catch block
            //try
            //{
                //make sure we are just accessing a session for code that follows
                using (Session epiSession = new Session("consignment", "consignment", "AppServerDC://app-server-1:9401", Session.LicenseType.Default))
                {
                    //set the plant to be the consignment one
                    epiSession.PlantID = "CONSIGN";

                    //connect to a Trans Order Receipt BO
                    TransOrderReceipt tor = new TransOrderReceipt(epiSession.ConnectionPool);

                    //connect to a Transfer Order Entry BO
                    TransferOrderEntry toe = new TransferOrderEntry(epiSession.ConnectionPool);

                    //call the warehouse bin bo
                    WhseBin wb = new WhseBin(epiSession.ConnectionPool);

                    //connect to a part bin search bo
                    PartBinSearch pbs = new PartBinSearch(epiSession.ConnectionPool);

                    //a list object for keeping track of which trays need to be inactivated
                    ArrayList WhseBinList = new ArrayList();

                    //a string that contains the from warehouse code
                    string FromWhseCode = "";

                    //get a dataset of the packs that aren't received
                    bool morePages;
                    TFShipHeadListDataSet listDS = tor.GetListBasicSearch(false, 0, 1, out morePages);

                    //loop through the packs that haven't been received
                    for (int i = 0; i < listDS.Tables["TFShipHeadList"].Rows.Count; i++)
                    {
                        //make sure this is a consignment order, and not a normal transfer
                        if (String.Compare(listDS.Tables["TFShipHeadList"].Rows[i]["ToPlant"].ToString(), "CONSIGN") == 0)
                        {
                            //get the order information
                            TransferOrderEntryDataSet toeDS = toe.GetByID(listDS.Tables["TFShipHeadList"].Rows[i]["ShortChar01"].ToString());

                            //check to see if the need by date is today or later
                            if (DateTime.Compare(Convert.ToDateTime(toeDS.Tables["TFOrdHed"].Rows[0]["Date01"]), DateTime.Today) == 0 ||
                                DateTime.Compare(Convert.ToDateTime(toeDS.Tables["TFOrdHed"].Rows[0]["Date01"]), DateTime.Today) < 0)
                            {
                                //get the data from the current pack
                                TransOrderReceiptDataSet torDS = tor.GetByID(Convert.ToInt32(listDS.Tables["TFShipHeadList"].Rows[i]["PackNum"]));

                                //set the from warehouse code string
                                FromWhseCode = torDS.Tables["PlantTran"].Rows[0]["FromWarehouseCode"].ToString();

                                //loop through all of the lines on the pack
                                for (int j = 0; j < torDS.Tables["PlantTran"].Rows.Count; j++)
                                {
                                    //get information about the order line
                                    bool morePages2;
                                    TransferOrderEntryDataSet ordLineDS = toe.GetRows("TFOrdNum = '" + torDS.Tables["TFShipDtl"].Rows[j]["TFOrdNum"] + "'", "", "TFOrdLine = '" + torDS.Tables["TFShipDtl"].Rows[j]["TFOrdLine"] + "'", 0, 1, out morePages2);

                                    //check to see if there is not a tray key
                                    if (String.IsNullOrEmpty(ordLineDS.Tables["TFOrdDtl"].Rows[0]["Character10"].ToString()))
                                    {
                                        //check to see if there is a replenish to bin selected, if not, move to bin 1
                                        if (!String.IsNullOrEmpty(ordLineDS.Tables["TFOrdDtl"].Rows[0]["ShortChar04"].ToString()))
                                        {
                                            //set the bin
                                            torDS.Tables["PlantTran"].Rows[j]["ReceiveToBinNum"] = ordLineDS.Tables["TFOrdDtl"].Rows[0]["ShortChar04"]; //the replenish to bin number
                                        }
                                        else
                                        {
                                            //set the bin to 1
                                            torDS.Tables["PlantTran"].Rows[j]["ReceiveToBinNum"] = 1;
                                        }
                                    }
                                    //if there is a tray key
                                    else
                                    {
                                        //get the information about the from warehouse bin
                                        bool morePages3;
                                        WhseBinDataSet wbDS_FromBin = wb.GetRows("WarehouseCode = '" + torDS.Tables["PlantTran"].Rows[j]["FromWarehouseCode"] + "' and BinNum = '" + torDS.Tables["PlantTran"].Rows[j]["BinNum"] + "'", "", 0, 1, out morePages3);

                                        //use the to warehouse and from bin to determine if it exists in the distributor plant
                                        bool morePages4;
                                        WhseBinDataSet wbDS_ToBin = wb.GetRows("WarehouseCode = '" + toeDS.Tables["TFOrdHed"].Rows[0]["ShortChar10"] + "' and BinNum = '" + torDS.Tables["PlantTran"].Rows[j]["BinNum"] + "'", "", 0, 1, out morePages4);

                                        //if there isn't a row, add a new bin to the warehouse
                                        if (wbDS_ToBin.Tables["WhseBin"].Rows.Count <= 0)
                                        {
                                            //get a new whsebin data set
                                            WhseBinDataSet wbDS_new = new WhseBinDataSet();

                                            //create a new bin
                                            wb.GetNewWhseBin(wbDS_new, toeDS.Tables["TFOrdHed"].Rows[0]["ShortChar10"].ToString());

                                            //set the bin number, bin description, warehouse zone, tray type, and tray revision
                                            wbDS_new.Tables["WhseBin"].Rows[0]["BinNum"] = wbDS_FromBin.Tables["WhseBin"].Rows[0]["BinNum"];
                                            wbDS_new.Tables["WhseBin"].Rows[0]["Description"] = wbDS_FromBin.Tables["WhseBin"].Rows[0]["Description"];
                                            wbDS_new.Tables["WhseBin"].Rows[0]["ZoneID"] = toeDS.Tables["TFOrdHed"].Rows[0]["ShortChar02"]; //whse zone from the order
                                            wbDS_new.Tables["WhseBin"].Rows[0]["Character01"] = wbDS_FromBin.Tables["WhseBin"].Rows[0]["Character01"];
                                            wbDS_new.Tables["WhseBin"].Rows[0]["Character06"] = wbDS_FromBin.Tables["WhseBin"].Rows[0]["Character06"];
                                            wbDS_new.Tables["WhseBin"].Rows[0]["Character04"] = torDS.Tables["PlantTran"].Rows[j]["PackNum"];

                                            //set the bin to be non-nettable 
                                            wbDS_new.Tables["WhseBin"].Rows[0]["NonNettable"] = false;

                                            //update the information
                                            wb.Update(wbDS_new);
                                        }
                                        //if there is a row, update the warehouse zone and last pack number
                                        else
                                        {
                                            wbDS_ToBin.Tables["WhseBin"].Rows[0]["ZoneID"] = toeDS.Tables["TFOrdHed"].Rows[0]["ShortChar02"];
                                            wbDS_ToBin.Tables["WhseBin"].Rows[0]["Character04"] = torDS.Tables["PlantTran"].Rows[j]["PackNum"];

                                            //update the information 
                                            wb.Update(wbDS_ToBin);
                                        }

                                        //check to see if the current plant tran row's bin isn't in the list
                                        if (!WhseBinList.Contains(wbDS_FromBin.Tables["WhseBin"].Rows[0]["BinNum"].ToString()))
                                        {
                                            //add the bin number to the list
                                            WhseBinList.Add(wbDS_FromBin.Tables["WhseBin"].Rows[0]["BinNum"].ToString());
                                        }

                                        //re-retrieve the data
                                        wbDS_ToBin = wb.GetRows("WarehouseCode = '" + toeDS.Tables["TFOrdHed"].Rows[0]["ShortChar10"] + "' and BinNum = '" + torDS.Tables["PlantTran"].Rows[j]["BinNum"] + "'", "", 0, 1, out morePages4);

                                        //check to see if the to warehouse bin is inactive
                                        if ((bool)wbDS_ToBin.Tables["WhseBin"].Rows[0]["InActive"] == true)
                                        {
                                            //activate it
                                            wbDS_ToBin.Tables["WhseBin"].Rows[0]["InActive"] = false;

                                            //update
                                            wb.Update(wbDS_ToBin);
                                        }

                                        //update the from bin's pack number
                                        wbDS_FromBin.Tables["WhseBin"].Rows[0]["Character04"] = torDS.Tables["PlantTran"].Rows[j]["PackNum"];
                                        wb.Update(wbDS_FromBin);

                                        //set the to warehouse code
                                        torDS.Tables["PlantTran"].Rows[j]["ReceiveToBinNum"] = wbDS_FromBin.Tables["WhseBin"].Rows[0]["BinNum"];
                                    }

                                    //change the receive to warehouse
                                    torDS.Tables["PlantTran"].Rows[j]["ReceiveToWhseCode"] = toeDS.Tables["TFOrdHed"].Rows[0]["ShortChar10"]; //the warehouse code

                                    //set the receive trans date to be right now
                                    torDS.Tables["PlantTran"].Rows[j]["RecTranDate"] = DateTime.Now;

                                    //update the reference to display the consignment order number
                                    torDS.Tables["PlantTran"].Rows[j]["PlntTranReference"] = "Consign Order: " + torDS.Tables["PlantTran"].Rows[j]["TFOrdNum"];

                                    //call the pre-update method
                                    bool requiresUserInput;
                                    tor.PreUpdate(torDS, out requiresUserInput);

                                    //call the update method
                                    tor.Update(torDS);
                                }
                            }
                        }
                    }

                    //check to see if there are any bins in the whse bin list
                    if (WhseBinList.Count > 0)
                    {
                        //loop through the bins in the whse bin list and inactivate them
                        for (int z = 0; z < WhseBinList.Count; z++)
                        {
                            //check to see if this is a valid whse bin
                            string errMsg;
                            pbs.CheckBin(FromWhseCode, WhseBinList[z].ToString(), out errMsg);

                            //if there isn't an error message
                            if (String.IsNullOrEmpty(errMsg))
                            {
                                //check to see if there are still contents in the bin
                                PartBinSearchDataSet pbsDS = pbs.GetBinContents(FromWhseCode, WhseBinList[z].ToString());

                                //if there aren't any rows
                                if (pbsDS.Tables["PartBinSearch"].Rows.Count <= 0)
                                {
                                    //get the dataset for each whse bin that needs to be inactivated
                                    WhseBinDataSet wbDS_InActive = wb.GetByID(FromWhseCode, WhseBinList[z].ToString());

                                    //set the inactive checkbox
                                    wbDS_InActive.Tables["WhseBin"].Rows[0]["InActive"] = true;

                                    //update 
                                    wb.Update(wbDS_InActive);
                                }
                            }
                        }
                    }
                }
            }
            /*catch
            {
                //create a new outlook application
                OutlookApp oApp = new OutlookApp();
                //create a new mail item
                MailItem mItem = oApp.CreateItem(OlItemType.olMailItem);
                //set the fields for the email
                mItem.To = "gkoutrelakos@amendia.com; solson@amendia.com";
                mItem.Subject = "Receive All Error";
                mItem.HTMLBody = "<html><body>There was an error while trying to run the Receive All script.</body></html>";
                //set the importance
                mItem.Importance = OlImportance.olImportanceHigh;
                //send the email
                ((_MailItem)mItem).Send();
            }*/
        }
    }
//}
