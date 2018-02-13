using System;
using System.Collections.Generic;
using System.Linq;
using AutoReceiver.Properties;
using Epicor.Mfg.BO;
using Epicor.Mfg.Core;
using NLog;

namespace AutoReceiver {
    public static class AutoReceiver {
        private static readonly Logger logger = LogManager.GetCurrentClassLogger();

        private static void Main() {
            try {
                logger.Info(
                    $"Connecting to: {Settings.Default.Server}:{Settings.Default.Port} Plant: {Settings.Default.Plant} User: {Settings.Default.User}");

                // Get unreceived packs for consignment plant
                var consignmentPacks = GetUnreceivedConsignmentPacks();

                /*
                 * Loop through all the packs!
                 * Step 1: Make sure destination bins exist for all lines.
                 * Step 2: Setup the line for receipt.
                 * Step 3: Add obsolete(source) bins to the deactivation list.
                 */
                var oldTrays = new List<Bin>();

                foreach (var consignmentPack in consignmentPacks)
                    ProcessPack(consignmentPack, oldTrays);

                // Inactivate all the obsolete bins.
                DeactiveBins(oldTrays);
            } catch (Exception ex) {
                logger.Error($"Something went seriously wrong. Processing cancelled. Error: {ex.Message}");
            } finally {
                ShutItDown();
                logger.Info("All Done!");
            }
        }

        private static void ShutItDown() {
            session.Dispose();
        }

        #region Pack Processing
        private static IEnumerable<TFShipHeadListDataSet.TFShipHeadListRow> GetUnreceivedConsignmentPacks() {
            logger.Info("Retrieving pending packs.");

            var packs = transOrderReceipt.GetListBasicSearch(false, 0, 1, out _);

            var consignmentPacks = packs.TFShipHeadList.Cast<TFShipHeadListDataSet.TFShipHeadListRow>()
                                        .Where(t => string.Equals(t.ToPlant, "CONSIGN", StringComparison.OrdinalIgnoreCase)).ToList();
            logger.Info($"Found {consignmentPacks.Count} packs.");

            return consignmentPacks;
        }

        /// <summary>
        ///     Create receipt bins, receive lines, update list of bins to deactivate.
        /// </summary>
        private static void ProcessPack(TFShipHeadListDataSet.TFShipHeadListRow consignmentPack, List<Bin> traysToInactivate) {
            try {
                logger.Info($"Processing pack: {consignmentPack.PackNum}");

                logger.Info($"Fetching order: {consignmentPack.ShortChar01}");

                var order         = transferOrderEntry.GetByID(consignmentPack.ShortChar01);
                var warehouseCode = order.TFOrdHed[0].ShortChar10;

                // We only want NeedBy dates before today
                var needBy = order.TFOrdHed[0].Date01;

                if (needBy > DateTime.Today)
                    return;

                logger.Info($"Fetching receipt: {consignmentPack.PackNum}");
                var receipt = transOrderReceipt.GetByID(consignmentPack.PackNum);

                logger.Info($"Found {receipt.PlantTran.Count} receipt lines.");

                foreach (var plantTran in receipt.PlantTran.Cast<TransOrderReceiptDataSet.PlantTranRow>())
                    ProcessLine(traysToInactivate, order, plantTran, warehouseCode, receipt);

                logger.Info("Pack complete!");
            } catch (Exception ex) {
                logger.Error($"Processing failed. Error: {ex.Message}");
            }
        }

        private static void ProcessLine(List<Bin> traysToInactivate, TransferOrderEntryDataSet order, TransOrderReceiptDataSet.PlantTranRow line,
                                        string warehouseCode, TransOrderReceiptDataSet receipt) {
            try {
                // Get the order line that matches this transaction to retrieve replenishment info.
                // Make sure this isn't case sensitive. The cases aren't consistent.
                var ordLineDS = order.TFOrdDtl.Cast<TransferOrderEntryDataSet.TFOrdDtlRow>().FirstOrDefault(
                    o => string.Equals(o.TFOrdNum,  line.TFOrdNum,  StringComparison.OrdinalIgnoreCase) &&
                         string.Equals(o.TFLineNum, line.TFLineNum, StringComparison.OrdinalIgnoreCase));
                
                if (ordLineDS==null)
                    throw new Exception("Unable to locate matching order line.");

                var trayKey          = ordLineDS.Character10;
                var replenishmentBin = ordLineDS.ShortChar04;

                logger.Info($"Line: {line.TFLineNum}. TrayKey: {trayKey} Replenishment bin: {replenishmentBin}");

                /*
                 * Set the receiving bin!
                 * If someone specified a 'trayKey' and replenishment bin, use that.
                 * Otherwise, ensure a duplicate of the source bin exists in the destination DB and use that.
                 */
                if (string.IsNullOrEmpty(trayKey))
                    line.ReceiveToBinNum = !string.IsNullOrEmpty(replenishmentBin) ? replenishmentBin : "1";
                else
                    CheckBinConfiguration(line, order, trayKey, replenishmentBin, traysToInactivate);

                // Update line!
                line.ReceiveToWhseCode = warehouseCode;
                line.RecTranDate       = DateTime.Now;
                line.PlntTranReference = $"Consign Order: {line.TFOrdNum}";

                transOrderReceipt.PreUpdate(receipt, out _);
                transOrderReceipt.Update(receipt);
            } catch (Exception ex) {
                logger.Error($"Error on line {line.TFLineNum}. Error: {ex.Message}");
            }
        }
        #endregion

        #region Bin Processing
        /// <summary>
        ///     Ensure a bin exists in the destination warehouse with the same ID as the source bin.
        ///     Update deactivation list.
        /// </summary>
        private static void CheckBinConfiguration(TransOrderReceiptDataSet.PlantTranRow plantTran, TransferOrderEntryDataSet order, string trayKey,
                                                  string replenishmentBin, List<Bin> traysToInactivate) {
            try {
                var toWarehouse           = order.TFOrdHed[0].ShortChar10;
                var toWarehouseZone       = order.TFOrdHed[0].ShortChar02;
                var fromWarehouse         = plantTran.FromWarehouseCode;
                var binNum                = plantTran.BinNum;
                var packNum               = plantTran.PackNum.ToString();
                var usingReplenishmentBin = true;

                // If trayKey is specified, we use the replenishment bin or fall back to "1".
                if (string.IsNullOrEmpty(trayKey)) {
                    binNum                = !string.IsNullOrEmpty(replenishmentBin) ? replenishmentBin : "1";
                    usingReplenishmentBin = false;
                }

                logger.Info($"Validating bins. Going from {fromWarehouse}/{binNum} to {toWarehouse}/{binNum}.");

                // Look up the source bin so we can copy properties to destination.
                var fromBinDS = whseBin.GetRows($@"WarehouseCode = '{fromWarehouse}' and BinNum = '{binNum}'", "", 0, 1, out _);

                UpsertBin(toWarehouse, binNum, toWarehouseZone, packNum, usingReplenishmentBin, fromBinDS);

                // Make sure the fromBin is in our deactivation list. (Unless it's a replenishment bin.)
                // It's going to be empty and obsolete after receipt.
                if (!usingReplenishmentBin && !traysToInactivate.Any(t => t.BinNum==binNum && t.WarehouseCode==fromWarehouse))
                    traysToInactivate.Add(new Bin {
                        BinNum        = binNum,
                        WarehouseCode = fromWarehouse
                    });

                plantTran.ReceiveToBinNum = binNum;
            } catch (Exception ex) {
                logger.Error($"Failed to validate bin configuration. Error: {ex.Message}");
            }
        }

        /// <summary>
        ///     Make sure bin exists with correct properties
        /// </summary>
        private static void UpsertBin(string warehouse, string binNum, string zoneID, string packNum, bool usingReplenishmentBin,
                                      WhseBinDataSet fromBinDS) {
            try {
                logger.Info($"Upserting {warehouse}/{binNum}.");

                // Maybe our bin already exists?
                var toBinDS = whseBin.GetRows($@"WarehouseCode = '{warehouse}' and BinNum = '{binNum}'", "", 0, 1, out _);

                // Find it? Create it? Just make sure it exists!
                if (toBinDS.WhseBin.Count > 0) {
                    logger.Info("Found bin!");
                } else {
                    logger.Info("Creating bin.");
                    toBinDS = new WhseBinDataSet();
                    whseBin.GetNewWhseBin(toBinDS, warehouse);
                    toBinDS.WhseBin[0].BinNum = binNum;
                    toBinDS.WhseBin[0].Description = binNum;
                }

                // Get those fields in sync!
                // Don't change the properties on a bin that may have already had them specified.
                if (!usingReplenishmentBin) {
                    logger.Info("Syncing bin fields.");
                    toBinDS.WhseBin[0].Description = fromBinDS.WhseBin[0].Description;
                    toBinDS.WhseBin[0].ZoneID      = zoneID;
                    toBinDS.WhseBin[0].Character04 = packNum;
                    toBinDS.WhseBin[0].Character01 = fromBinDS.WhseBin[0].Character01; // Tray Type
                    toBinDS.WhseBin[0].Character06 = fromBinDS.WhseBin[0].Character06; // Tray Revision
                    toBinDS.WhseBin[0].NonNettable = false;
                }

                // The bin has to be active or we can't receieve into it.
                toBinDS.WhseBin[0].InActive = false;

                whseBin.Update(toBinDS);
            } catch (Exception ex) {
                logger.Error($"Error upserting bin. Error: {ex.Message}");
            }
        }

        /// <summary>
        ///     Go through a list of bins and deactivate them all.
        ///     Bins with stuff in them will be ignored.
        /// </summary>
        private static void DeactiveBins(List<Bin> bins) {
            logger.Info($"Deactivating {bins.Count} bins.");

            foreach (var bin in bins)
                DeactivateBin(bin);

            logger.Info("Done with bins!");
        }

        private static void DeactivateBin(Bin bin) {
            try {
                logger.Info($"Deactivating bin {bin}.");

                // Is this a valid bin?
                partBinSearch.CheckBin(bin.WarehouseCode, bin.BinNum, out var errMsg);

                // Something went wrong!
                if (!string.IsNullOrEmpty(errMsg))
                    throw new Exception(errMsg);

                // Is there anything in the bin?
                if (partBinSearch.GetBinContents(bin.WarehouseCode, bin.BinNum).PartBinSearch.Count > 0) {
                    logger.Info("Bin isn't empty. Deactivation cancelled.");
                    return;
                }

                // Inactivate our bin!
                var wbDS_InActive = whseBin.GetByID(bin.WarehouseCode, bin.BinNum);
                wbDS_InActive.WhseBin[0].InActive = true;
                whseBin.Update(wbDS_InActive);
            } catch (Exception ex) {
                logger.Error($"Failed to deactivate bin. Error: {ex.Message}");
            }
        }
        #endregion

        #region Epicor Properties
        private static Session _session;
        private static Session session =>
            _session ??
            (_session = new Session(Settings.Default.User, Settings.Default.Pass, $"AppServerDC://{Settings.Default.Server}:{Settings.Default.Port}",
                                    Session.LicenseType.Default) {
                    PlantID = Settings.Default.Plant
                });
        private static WhseBin            whseBin            => new WhseBin(session.ConnectionPool);
        private static PartBinSearch      partBinSearch      => new PartBinSearch(session.ConnectionPool);
        private static TransOrderReceipt  transOrderReceipt  => new TransOrderReceipt(session.ConnectionPool);
        private static TransferOrderEntry transferOrderEntry => new TransferOrderEntry(session.ConnectionPool);
        #endregion
    }

    internal class Bin {
        public string BinNum;
        public string WarehouseCode;
    }
}