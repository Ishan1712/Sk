import * as React from 'react';
import styles from './QuotationResult.module.scss';
import { IQuotationResultProps } from './IQuotationResultProps';
import { useState, useEffect } from 'react';
import { Web } from 'sp-pnp-js';

interface QuotationData {
  id: number;
  quotationDate: string;
  serialNumber: string;
  totalAmount: number;
  totalRate: number;
  totalWeight: number;
  status: string;
  revisionNumber: number;
  lossreason?:string;
}

const QuotationResult: React.FC<IQuotationResultProps> = ({ userDisplayName }) => {
  const [quotations, setQuotations] = useState<QuotationData[]>([]);
  const [showReasonModal, setShowReasonModal] = useState(false);
  const [selectedReason, setSelectedReason] = useState('');
  const [currentQuotation, setCurrentQuotation] = useState<QuotationData | null>(null);
  const reasonOptions = ["Delivery", "Budget", "Material Specification"];

  const fetchHighestRevisionRecord = async (rfqNumber: string, web: Web): Promise<QuotationData | null> => {
    try {
      // Fetch records from QuotationList
      const quotationItems = await web.lists
        .getByTitle("QuotationList")
        .items.filter(`RFQSerialNumber eq '${rfqNumber}' and Status eq 'Waiting'`)
        .select("ID", "RFQSerialNumber", "RevisionNumber", "QuotationDate", "TotalAmount", "TotalRate", "TotalWeight", "Status","LossReason")
        .get();
  
      // Check for a record in QuotationList with RevisionNumber 0 and Status 'Waiting'
      const zeroRevisionItem = quotationItems.find((item: { 
        RevisionNumber?: string; 
        Status?: string; 
      }) => parseInt(item.RevisionNumber || "0", 10) === 0 && item.Status === "Waiting");
  
      if (zeroRevisionItem) {
        // If a valid revision 0 item exists in QuotationList, return it directly
        return {
          id: zeroRevisionItem.ID,
          serialNumber: zeroRevisionItem.RFQSerialNumber || "",
          quotationDate: zeroRevisionItem.QuotationDate || "",
          totalAmount: parseFloat(zeroRevisionItem.TotalAmount || "0"),
          totalRate: parseFloat(zeroRevisionItem.TotalRate || "0"),
          totalWeight: parseFloat(zeroRevisionItem.TotalWeight || "0"),
          status: zeroRevisionItem.Status || "",
          revisionNumber: 0,
          lossreason:zeroRevisionItem.LossReason || "",
        };
      }
  
      // Fetch records from QuotationRevision if no valid revision 0 is found
      const revisionItems = await web.lists
        .getByTitle("QuotationRevision")
        .items.filter(`RFQSerialNumber eq '${rfqNumber}' and Statuss eq 'Waiting'`)
        .select("ID", "RFQSerialNumber", "RevisionNumber", "RevisionDate", "TotalAmount", "TotalRate", "TotalWeight", "Statuss")
        .get();
  
      // Combine records and map to uniform structure
      const allItems = [
        ...quotationItems.map((item: any) => ({
          ...item,
          Status: item.Status, // Status is directly from QuotationList
        })),
        ...revisionItems.map((item: any) => ({
          ...item,
          Status: item.Statuss, // Map Statuss from QuotationRevision to Status
        })),
      ];
  
      // Find the highest revision number
      let highestRecord: QuotationData | null = null;
      let highestRevisionNumber = 0;
  
      allItems.forEach((item) => {
        const revisionNumber = parseInt(item.RevisionNumber || "0", 10);
        if (revisionNumber > highestRevisionNumber) {
          highestRevisionNumber = revisionNumber;
          highestRecord = {
            id: item.ID,
            serialNumber: item.RFQSerialNumber || "",
            quotationDate: item.QuotationDate || item.RevisionDate || "",
            totalAmount: parseFloat(item.TotalAmount || "0"),
            totalRate: parseFloat(item.TotalRate || "0"),
            totalWeight: parseFloat(item.TotalWeight || "0"),
            status: item.Status || "",
            revisionNumber: revisionNumber,
            // lossreason: item.LossReason || "",
          };
        }
      });
  
      return highestRecord;
    } catch (error) {
      console.error(`Error fetching highest revision record for RFQ ${rfqNumber}:`, error);
      return null;
    }
  };
  
  const fetchQuotations = async () => {
    try {
      const web = new Web("https://skgroupenginering.sharepoint.com/sites/SalesManagement");

      // Fetch RFQ Serial Numbers from both lists
      const quotationItems = await web.lists
        .getByTitle('QuotationList')
        .items.filter("Status eq 'Waiting'")
        .select('RFQSerialNumber')
        .get();

      const revisionItems = await web.lists
        .getByTitle('QuotationRevision')
        .items.filter("Status eq 'Waiting'")
        .select('RFQSerialNumber')
        .get();

      // Combine and deduplicate RFQ Serial Numbers
      const rfqSerialNumbers: string[] = [];
      const rfqSet = new Set<string>();

      [...quotationItems, ...revisionItems].forEach((item) => {
        const rfqSerialNumber = item.RFQSerialNumber;
        if (rfqSerialNumber && !rfqSet.has(rfqSerialNumber)) {
          rfqSet.add(rfqSerialNumber);
          rfqSerialNumbers.push(rfqSerialNumber);
        }
      });

      // Fetch the highest revision record for each RFQ Serial Number
      const results = await Promise.all(
        rfqSerialNumbers.map(async (rfqNumber) => {
          return await fetchHighestRevisionRecord(rfqNumber, web);
        })
      );

      setQuotations(results.filter((q): q is QuotationData => q !== null));
    } catch (error) {
      console.error('Error fetching quotations:', error);
    }
  };

  const updateQuotationStatus = async (
    quotation: QuotationData,
    newStatus: string,
    reason?: string
  ) => {
    try {
      if (!quotation || !quotation.serialNumber) {
        throw new Error("Invalid quotation data: 'quotation' or 'serialNumber' is missing.");
      }

      const web = new Web("https://skgroupenginering.sharepoint.com/sites/SalesManagement");

      console.log(`Updating status for RFQSerialNumber: ${quotation.serialNumber} to "${newStatus}"`);

      // Step 1: Update QuotationList records based on RFQSerialNumber
      const quotationItems = await web.lists
        .getByTitle("QuotationList")
        .items.filter(`RFQSerialNumber eq '${quotation.serialNumber}'`)
        .select("ID")
        .get();

      if (quotationItems.length > 0) {
        await Promise.all(
          quotationItems.map((item) =>
            web.lists
              .getByTitle("QuotationList")
              .items.getById(item.ID)
              .update({
                Status: newStatus,
                LossReason: reason || "No reason provided",
              })
          )
        );
        console.log("QuotationList records updated successfully.");
      } else {
        console.warn(`No records found in QuotationList for RFQSerialNumber: ${quotation.serialNumber}`);
      }

      // Step 2: Update QuotationRevision records based on RFQSerialNumber
      const revisionItems = await web.lists
        .getByTitle("QuotationRevision")
        .items.filter(`RFQSerialNumber eq '${quotation.serialNumber}'`)
        .select("ID")
        .get();

      if (revisionItems.length > 0) {
        await Promise.all(
          revisionItems.map((item) =>
            web.lists
              .getByTitle("QuotationRevision")
              .items.getById(item.ID)
              .update({
                Statuss: newStatus, // Statuss field in QuotationRevision
              })
          )
        );
        console.log("QuotationRevision records updated successfully.");
      } else {
        console.warn(`No records found in QuotationRevision for RFQSerialNumber: ${quotation.serialNumber}`);
      }

      // Step 3: Update RFQList records based on RFQNumber
      const rfqItems = await web.lists
        .getByTitle("RFQList")
        .items.filter(`RFQNumber eq '${quotation.serialNumber}'`)
        .select("ID")
        .get();

      if (rfqItems.length > 0) {
        await Promise.all(
          rfqItems.map((item) =>
            web.lists
              .getByTitle("RFQList")
              .items.getById(item.ID)
              .update({
                Status: newStatus,
              })
          )
        );
        console.log("RFQList records updated successfully.");
      } else {
        console.warn(`No records found in RFQList for RFQNumber: ${quotation.serialNumber}`);
      }

      // Step 4: Update the frontend table to reflect the changes
      setQuotations((prevQuotations) =>
        prevQuotations.map((q) =>
          q.serialNumber === quotation.serialNumber
            ? { ...q, status: newStatus, lossreason: reason || "No reason provided" }
            : q
        )
      );

      alert(`Status updated in QuotationList, QuotationRevision, and RFQList to "${newStatus}" successfully.`);
    } catch (error) {
      console.error("Error updating quotation or RFQ status:", error);
      alert("An error occurred while updating the status. Please check the console for more details.");
    }
  };

  
  const handleLossClick = (quotation: QuotationData) => {
    setCurrentQuotation(quotation);
    setShowReasonModal(true);
  };


  const handleReasonSubmit = () => {
    if (!selectedReason) {
      alert("Please select a reason before submitting.");
      return;
    }
  
    if (currentQuotation) {
      updateQuotationStatus(currentQuotation, "Loss", selectedReason); // Pass LossReason to updateQuotationStatus
    }
  
    setShowReasonModal(false);
    setSelectedReason('');
    setCurrentQuotation(null);
  };
  

  useEffect(() => {
    fetchQuotations();
  }, []);



  return (
    <section className={styles.quotationresultscreen}>
      <div className={styles.quotationresultscreen}>
        <h2>Quotation Result Screen</h2>
        <div className={styles.scrollableContainer}>
          <table className={styles.table}>
            <thead>
              <tr>
                <th>Quotation Date</th>
                <th>RFQ Serial Number</th>
                <th>Total Amount</th>
                <th>Total Rate</th>
                <th>Total Weight</th>
                <th>Status</th>
                <th>Reason</th>
                <th>Actions</th>
              </tr>
            </thead>
            <tbody>
              {quotations.map((quotation) => (
                <tr key={quotation.id}>
                  <td>{quotation.quotationDate}</td>
                  <td>{quotation.serialNumber}</td>
                  <td>{quotation.totalAmount}</td>
                  <td>{quotation.totalRate}</td>
                  <td>{quotation.totalWeight}</td>
                  <td>{quotation.status}</td>
                  <td>{quotation.lossreason || "N/A"}</td> 
                  <td className={styles.actionsColumn}>
                    <div className={styles.buttonGroup}>
                      <button
                        className={styles.wonButton}
                        onClick={() => updateQuotationStatus(quotation, "Won")}
                      >
                        Won
                      </button>
                      <button
                        className={styles.lossButton}
                        onClick={() => handleLossClick(quotation)}
                      >
                        Loss
                      </button>
                      <button
                        className={styles.reviseButton}
                        onClick={() => updateQuotationStatus(quotation, "Revised")}
                      >
                        Revise
                      </button>
                    </div>
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      </div>

      {showReasonModal && (
        <div className={styles.modal}>
          <div className={styles.modalContent}>
            <h3>Select the Appropriate Reason</h3>
            <select
              value={selectedReason}
              onChange={(e) => setSelectedReason(e.target.value)}
              className={styles.dropdown}
            >
              <option value="">-- Select Reason --</option>
              {reasonOptions.map((reason) => (
                <option key={reason} value={reason}>
                  {reason}
                </option>
              ))}
            </select>
            <div className={styles.modalActions}>
              <button onClick={handleReasonSubmit} className={styles.submitButton}>
                Submit
              </button>
              <button
                onClick={() => {
                  setShowReasonModal(false);
                  setSelectedReason('');
                  setCurrentQuotation(null);
                }}
                className={styles.cancelButton}
              >
                Cancel
              </button>
            </div>
          </div>
        </div>
      )}

    </section>
  );
};

export default QuotationResult;