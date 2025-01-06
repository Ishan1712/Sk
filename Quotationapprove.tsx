import * as React from 'react';
import styles from './Quotationapprove.module.scss';
import { IQuotationapproveProps } from './IQuotationapproveProps';
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
  reason?: string; 
  approvalDate?: string;
}

const fetchHighestRevisionRecord = async (rfqNumber: string, web: Web): Promise<QuotationData | null> => {
  try {
    // Fetch records from QuotationList
    const quotationItems = await web.lists
      .getByTitle("QuotationList")
      .items.filter(`RFQSerialNumber eq '${rfqNumber}'`)
      .select("ID", "RFQSerialNumber", "RevisionNumber", "QuotationDate", "TotalAmount", "TotalRate", "TotalWeight", "Status", "Reason", "ApprovalDate")
      .get();

    // Fetch records from QuotationRevision
    const revisionItems = await web.lists
      .getByTitle("QuotationRevision")
      .items.filter(`RFQSerialNumber eq '${rfqNumber}'`)
      .select("ID", "RFQSerialNumber", "RevisionNumber", "RevisionDate", "TotalAmount", "TotalRate", "TotalWeight", "Statuss", "Reason", "ApprovalDate")
      .get();

    // Find the highest revision number across both lists
    let highestRevisionNumber = 0;
    let highestSource = "QuotationList"; // Track the source of the highest revision

    quotationItems.forEach((item :any) => {
      const revisionNumber = parseInt(item.RevisionNumber || "0", 10);
      if (revisionNumber > highestRevisionNumber) {
        highestRevisionNumber = revisionNumber;
        highestSource = "QuotationList";
      }
    });

    revisionItems.forEach((item:any) => {
      const revisionNumber = parseInt(item.RevisionNumber || "0", 10);
      if (revisionNumber > highestRevisionNumber) {
        highestRevisionNumber = revisionNumber;
        highestSource = "QuotationRevision";
      }
    });

    if (highestSource === "QuotationList") {
      // Fetch the record with the highest revision number from QuotationList
      const highestRevisionItem = quotationItems.find((item: {
        RevisionNumber?: string;
      }) => parseInt(item.RevisionNumber || "0", 10) === highestRevisionNumber);

      if (highestRevisionItem) {
        return {
          id: highestRevisionItem.ID,
          serialNumber: highestRevisionItem.RFQSerialNumber || "",
          quotationDate: highestRevisionItem.QuotationDate || "",
          totalAmount: parseFloat(highestRevisionItem.TotalAmount || "0"),
          totalRate: parseFloat(highestRevisionItem.TotalRate || "0"),
          totalWeight: parseFloat(highestRevisionItem.TotalWeight || "0"),
          status: highestRevisionItem.Status || "",
          revisionNumber: highestRevisionNumber,
          reason: highestRevisionItem.Reason || "",
          approvalDate: highestRevisionItem.ApprovalDate || "",
        };
      }
    } else if (highestSource === "QuotationRevision") {
      // Fetch the record with the highest revision number from QuotationRevision
      const highestRevisionItem = revisionItems.find((item: {
        RevisionNumber?: string;
      }) => parseInt(item.RevisionNumber || "0", 10) === highestRevisionNumber);

      if (highestRevisionItem) {
        return {
          id: highestRevisionItem.ID,
          serialNumber: highestRevisionItem.RFQSerialNumber || "",
          quotationDate: highestRevisionItem.RevisionDate || "",
          totalAmount: parseFloat(highestRevisionItem.TotalAmount || "0"),
          totalRate: parseFloat(highestRevisionItem.TotalRate || "0"),
          totalWeight: parseFloat(highestRevisionItem.TotalWeight || "0"),
          status: highestRevisionItem.Statuss || "",
          revisionNumber: highestRevisionNumber,
          reason: highestRevisionItem.Reason || "",
          approvalDate: highestRevisionItem.ApprovalDate || "",
        };
      }
    }

    return null;
  } catch (error) {
    console.error(`Error fetching highest revision record for RFQ ${rfqNumber}:`, error);
    return null;
  }
};


const Quotationapprove: React.FC<IQuotationapproveProps> = ({ userDisplayName }) => {
  const [quotations, setQuotations] = useState<QuotationData[]>([]);
  const [showConfirmModal, setShowConfirmModal] = useState(false);
  const [showRejectModal, setShowRejectModal] = useState(false);
  const [currentQuotation, setCurrentQuotation] = useState<QuotationData | null>(null);

  
  useEffect(() => {
    const fetchAllHighestRevisions = async () => {
      try {
        const web = new Web("https://skgroupenginering.sharepoint.com/sites/SalesManagement");
  
        // Fetch RFQ Serial Numbers from QuotationList where Status is 'WorkingDone'
        const quotationItems = await web.lists
          .getByTitle("QuotationList")
          .items.filter("Status eq 'WorkingDone'")
          .select("RFQSerialNumber")
          .get();
  
        // Fetch RFQ Serial Numbers from QuotationRevision where Statuss is 'WorkingDone'
        const revisionItems = await web.lists
          .getByTitle("QuotationRevision")
          .items.filter("Statuss eq 'WorkingDone'")
          .select("RFQSerialNumber")
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
        const updatedQuotations = await Promise.all(
          rfqSerialNumbers.map(async (rfqNumber) => {
            return await fetchHighestRevisionRecord(rfqNumber, web);
          })
        );
  
        // Update the quotations state
        setQuotations(updatedQuotations.filter((q): q is QuotationData => q !== null));
      } catch (error) {
        console.error("Error fetching quotations:", error);
      }
    };
  
    fetchAllHighestRevisions();
  }, []);
  
  
  const approveQuotation = async (quotation: QuotationData) => {
    try {
      const web = new Web("https://skgroupenginering.sharepoint.com/sites/SalesManagement");
      const currentDate = new Date().toISOString(); // Get the current date in ISO format
  
      // Update the QuotationList item if it exists
      const quotationListItems = await web.lists
        .getByTitle("QuotationList")
        .items.filter(`RFQSerialNumber eq '${quotation.serialNumber}'`)
        .get();
  
      if (quotationListItems.length > 0) {
        await Promise.all(
          quotationListItems.map((item: { ID: number }) =>
            web.lists
              .getByTitle("QuotationList")
              .items.getById(item.ID)
              .update({
                Status: "Approved",
                ApprovalDate: currentDate,
              })
          )
        );
      }
  
      // Update the QuotationRevision item if it exists
      const quotationRevisionItems = await web.lists
        .getByTitle("QuotationRevision")
        .items.filter(`RFQSerialNumber eq '${quotation.serialNumber}'`)
        .get();
  
      if (quotationRevisionItems.length > 0) {
        await Promise.all(
          quotationRevisionItems.map((item: { ID: number }) =>
            web.lists
              .getByTitle("QuotationRevision")
              .items.getById(item.ID)
              .update({
                Statuss: "Approved",
                ApprovalDate: currentDate,
              })
          )
        );
      }
  
      // Update the RFQList item if it exists
      const rfqListItems = await web.lists
        .getByTitle("RFQList")
        .items.filter(`RFQNumber eq '${quotation.serialNumber}'`)
        .get();
  
      if (rfqListItems.length > 0) {
        await Promise.all(
          rfqListItems.map((item: { ID: number }) =>
            web.lists
              .getByTitle("RFQList")
              .items.getById(item.ID)
              .update({
                Status: "Approved",
              })
          )
        );
      }
  
      alert("Quotation and RFQ statuses updated to Approved.");
      setQuotations((prevQuotations) =>
        prevQuotations.filter((q) => q.id !== quotation.id)
      );
    } catch (error) {
      console.error("Error updating quotation or RFQ status:", error);
    }
  };
  
  
  const handleSubmit = (quotation: QuotationData) => {
    setShowConfirmModal(true); // Show confirmation modal
    setCurrentQuotation(quotation);
  };

  const confirmSubmit = async () => {
    if (currentQuotation) {
      await approveQuotation(currentQuotation); // Approve the selected quotation
      setShowConfirmModal(false); // Close the modal
      setCurrentQuotation(null);
    }
  };
  
  const rejectQuotation = async (quotation: QuotationData) => {
    try {
      const web = new Web("https://skgroupenginering.sharepoint.com/sites/SalesManagement");
  
      // Update the QuotationList item if it exists
      const quotationListItems = await web.lists
        .getByTitle("QuotationList")
        .items.filter(`RFQSerialNumber eq '${quotation.serialNumber}'`)
        .get();
  
      if (quotationListItems.length > 0) {
        await Promise.all(
          quotationListItems.map((item: { ID: number }) =>
            web.lists
              .getByTitle("QuotationList")
              .items.getById(item.ID)
              .update({
                Status: "Todo",
                Reason: quotation.reason || "No reason provided",
              })
          )
        );
      }
  
      // Update the QuotationRevision item if it exists
      const quotationRevisionItems = await web.lists
        .getByTitle("QuotationRevision")
        .items.filter(`RFQSerialNumber eq '${quotation.serialNumber}'`)
        .get();
  
      if (quotationRevisionItems.length > 0) {
        await Promise.all(
          quotationRevisionItems.map((item: { ID: number }) =>
            web.lists
              .getByTitle("QuotationRevision")
              .items.getById(item.ID)
              .update({
                Statuss: "Todo",
                Reason: quotation.reason || "No reason provided",
              })
          )
        );
      }
  
      // Update the RFQList item if it exists
      const rfqListItems = await web.lists
        .getByTitle("RFQList")
        .items.filter(`RFQNumber eq '${quotation.serialNumber}'`)
        .get();
  
      if (rfqListItems.length > 0) {
        await Promise.all(
          rfqListItems.map((item: { ID: number }) =>
            web.lists
              .getByTitle("RFQList")
              .items.getById(item.ID)
              .update({
                Status: "Todo",
              })
          )
        );
      }
  
      alert("Quotation and RFQ statuses updated to Rejected.");
      setQuotations((prevQuotations) =>
        prevQuotations.filter((q) => q.id !== quotation.id)
      );
    } catch (error) {
      console.error("Error updating quotation or RFQ status:", error);
    }
  };
  
  
  const handleReject = (quotation: QuotationData) => {
    setShowRejectModal(true);
    setCurrentQuotation(quotation);
  };

  const confirmReject = async () => {
    if (currentQuotation) {
      await rejectQuotation(currentQuotation);
      setShowRejectModal(false);
      setCurrentQuotation(null);
    }
  };

  return (
    <div className={styles.quorevisionscreen}>
      <h2>Quotation Approval Screen </h2>
    <div className={styles.scrollableContainer}>
      <table className={styles.table}>
        <thead>
          <tr>
            <th>Quotation Date</th>
            <th>RFQ Serial Number</th>
            <th>Total Amount</th>
            {/* <th>Total Rate</th> */}
            <th>Total Weight</th>
            <th>Status</th>
            <th>Reason</th>
            <th>Approval Date</th>
            <th>Actions</th>
          </tr>
        </thead>
        <tbody>
          {quotations.map((quotation) => (
            <tr key={quotation.id}>
              <td>{quotation.quotationDate}</td>
              <td>{quotation.serialNumber}</td>
              <td>{quotation.totalAmount}</td>
              {/* <td>{quotation.totalRate}</td> */}
              <td>{quotation.totalWeight}</td>
              <td>{quotation.status}</td>
              <td>{quotation.reason || "N/A"}</td> 
              <td>{quotation.approvalDate || "N/A"}</td>
              <td className={styles.actionsColumn}>
                <div className={styles.buttonGroup}>
                  <button className={styles.submitButton} onClick={() => handleSubmit(quotation)}>Approve</button>
                  <button className={styles.deleteButton} onClick={() => handleReject(quotation)}>Reject</button>
                </div>
              </td>
            </tr>
          ))}
        </tbody>
      </table>
      </div>

      {/* Confirmation Modal for Submission */}
      {showConfirmModal && (
        <div className={styles.modal}>
          <div className={styles.modalContent}>
            <p className={styles.modalMessage}>Are you sure you want to approve this data?</p>
            <div className={styles.buttonGroup}>
              <button onClick={confirmSubmit} className={styles.yesButton}>Yes</button>
              <button onClick={() => setShowConfirmModal(false)} className={styles.noButton}>No</button>
            </div>
          </div>
        </div>
      )}

      {showRejectModal && (
          <div className={styles.modal}>
             <div className={styles.modalContent}>
                <p className={styles.modalMessage}>Why are you rejecting this quotation?</p>
                <textarea
  className={styles.textArea}
  placeholder="Enter rejection reason"
  value={currentQuotation?.reason || ""}
  onChange={(e) => {
    const updatedReason = e.target.value;

    setQuotations((prevQuotations) =>
      prevQuotations.map((quotation) =>
        quotation.id === currentQuotation?.id
          ? { ...quotation, reason: updatedReason }
          : quotation
      )
    );

    setCurrentQuotation((prev) =>
      prev ? { ...prev, reason: updatedReason } : null
    );
  }}
/>
          <div className={styles.buttonGroup}>
            <button onClick={confirmReject} className={styles.yesButton}>Yes</button>
            <button onClick={() => setShowRejectModal(false)} className={styles.noButton}>No</button>
          </div>
          </div>
        </div>
)}

    </div>
  );
};
export default Quotationapprove;