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
}

const QuotationResult: React.FC<IQuotationResultProps> = ({ userDisplayName }) => {
  const [quotations, setQuotations] = useState<QuotationData[]>([]);

  // Fetch records with "Waiting" status from SharePoint
  const fetchQuotations = async (): Promise<void> => {
    try {
      const web = new Web("https://skgroupenginering.sharepoint.com/sites/SalesManagement")
      const items = await web.lists
        .getByTitle("QuotationList")
        .items.select("ID", "QuotationDate", "RFQSerialNumber", "TotalAmount", "TotalRate", "TotalWeight", "Status")
        .filter("Status eq 'Waiting'")
        .get();

      const fetchedQuotations = items.map((item) => ({
        id: item.ID,
        quotationDate: item.QuotationDate || "",
        serialNumber: item.RFQSerialNumber || "",
        totalAmount: item.TotalAmount || 0,
        totalRate: item.TotalRate || 0,
        totalWeight: item.TotalWeight || 0,
        status: item.Status || "Working Done",
      }));

      setQuotations(fetchedQuotations);
    } catch (error) {
      console.error("Error fetching quotations:", error);
    }
  };


  const updateQuotationStatus = async (quotation: QuotationData, newStatus: string) => {
    // Show confirmation popup
    const confirmed = window.confirm(`Are you sure you want to update the status to '${newStatus}'?`);
    if (!confirmed) return; // Exit if the user cancels

    try {
      const web = new Web("https://skgroupenginering.sharepoint.com/sites/SalesManagement") // Specify your SharePoint site URL here.

      // Update status in QuotationList
      await web.lists
        .getByTitle("QuotationList")
        .items.getById(quotation.id)
        .update({
          Status: newStatus,
        });

      // Update status in RFQList
      const rfqItems = await web.lists
        .getByTitle("RFQList")
        .items.filter(`RFQNumber eq '${quotation.serialNumber}'`)
        .get();

      if (rfqItems.length > 0) {
        const rfqItemId = rfqItems[0].ID; // Assume the first matching RFQ is the correct one.
        await web.lists
          .getByTitle("RFQList")
          .items.getById(rfqItemId)
          .update({
            Status: newStatus,
          });
      }

      alert(`Quotation and RFQ status updated to ${newStatus}.`);
      fetchQuotations(); // Refresh the list to update the status
    } catch (error) {
      console.error(`Error updating quotation and RFQ status to ${newStatus}:`, error);
    }
  };
  useEffect(() => {
    fetchQuotations(); // Fetch records when the component mounts
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
                        onClick={() => updateQuotationStatus(quotation, "Lost")}
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
    </section>
  );
};

export default QuotationResult;
