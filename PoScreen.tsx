import * as React from 'react';
import { useState, useEffect ,useMemo } from 'react';
import styles from './PoScreen.module.scss';
import { IPoScreenProps } from './IPoScreenProps';
import { Web } from 'sp-pnp-js';

interface IPORecord {
  id: number;
  serialNumber: string;
  rfqNumber: string;
  totalRate: number;
  totalWeight: number;
  totalAmount: number;
  poNumber?: string;
  poDate?: string;
  rateByCustomer?: string;
  pdfUrl?: string;
  uploading?: boolean;
  revisionNumber?: number;
}

const fetchHighestRevisionRecord = async (rfqNumber: string, web: Web): Promise<IPORecord | null> => {
  try {
    const quotationItems = await web.lists
      .getByTitle("QuotationList")
      .items.filter(`RFQSerialNumber eq '${rfqNumber}'`)
      .select("ID", "RFQSerialNumber", "RevisionNumber", "TotalAmount", "TotalRate", "TotalWeight", "Status")
      .get();

    const revisionItems = await web.lists
      .getByTitle("QuotationRevision")
      .items.filter(`RFQSerialNumber eq '${rfqNumber}'`)
      .select("ID", "RFQSerialNumber", "RevisionNumber", "RevisionDate", "TotalAmount", "TotalRate", "TotalWeight", "Statuss")
      .get();

    let highestRevisionNumber = 0;
    let highestSource = "QuotationList";

    quotationItems.forEach((item) => {
      const revisionNumber = parseInt(item.RevisionNumber || "0", 10);
      if (revisionNumber > highestRevisionNumber) {
        highestRevisionNumber = revisionNumber;
        highestSource = "QuotationList";
      }
    });

    revisionItems.forEach((item) => {
      const revisionNumber = parseInt(item.RevisionNumber || "0", 10);
      if (revisionNumber > highestRevisionNumber) {
        highestRevisionNumber = revisionNumber;
        highestSource = "QuotationRevision";
      }
    });

    let highestRevisionItem = null;

    if (highestSource === "QuotationList") {
      highestRevisionItem = quotationItems.find(
        (item) => parseInt(item.RevisionNumber || "0", 10) === highestRevisionNumber
      );
    } else {
      highestRevisionItem = revisionItems.find(
        (item) => parseInt(item.RevisionNumber || "0", 10) === highestRevisionNumber
      );
    }

    if (highestRevisionItem) {
      return {
        id: highestRevisionItem.ID,
        serialNumber: highestRevisionItem.RFQSerialNumber || "",
        rfqNumber: highestRevisionItem.RFQSerialNumber || "",
        totalRate: parseFloat(highestRevisionItem.TotalRate || "0"),
        totalWeight: parseFloat(highestRevisionItem.TotalWeight || "0"),
        totalAmount: parseFloat(highestRevisionItem.TotalAmount || "0"),
        revisionNumber: highestRevisionNumber,
      };
    }

    return null;
  } catch (error) {
    console.error(`Error fetching highest revision record for RFQ ${rfqNumber}:`, error);
    return null;
  }
};

const POscreen: React.FC<IPoScreenProps> = (props) => {
  const [records, setRecords] = useState<IPORecord[]>([]);
  const [recordDetails, setRecordDetails] = useState<IPORecord | null>(null);
  const [showPOForm, setShowPOForm] = useState(false);
  const [poNumber, setPoNumber] = useState('');
  const [poDate, setPoDate] = useState('');
  const [rateByCustomer, setRateByCustomer] = useState('');

  useEffect(() => {
    const fetchAllHighestRevisions = async () => {
      try {
        const web = new Web("https://skgroupenginering.sharepoint.com/sites/SalesManagement");

        const quotationItems = await web.lists
          .getByTitle("QuotationList")
          .items.filter("Status eq 'Won'")
          .select("RFQSerialNumber")
          .get();

        const revisionItems = await web.lists
          .getByTitle("QuotationRevision")
          .items.filter("Statuss eq 'Won'")
          .select("RFQSerialNumber")
          .get();

        const rfqSerialNumbers = [
          ...quotationItems.map((item) => item.RFQSerialNumber),
          ...revisionItems.map((item) => item.RFQSerialNumber),
        ].filter((value, index, self) => self.indexOf(value) === index);

        const updatedRecords = await Promise.all(
          rfqSerialNumbers.map(async (rfqNumber) => {
            return await fetchHighestRevisionRecord(rfqNumber, web);
          })
        );

        setRecords(updatedRecords.filter((record): record is IPORecord => record !== null));
      } catch (error) {
        console.error("Error fetching highest revisions:", error);
      }
    };

    fetchAllHighestRevisions();
  }, []);

  const handlePOFormChange = (event: React.ChangeEvent<HTMLInputElement>) => {
    const { name, value } = event.target;
    if (name === 'poNumber') setPoNumber(value);
    if (name === 'poDate') setPoDate(value);
    if (name === 'rateByCustomer') setRateByCustomer(value);
  };

  const savePO = async () => {
    if (!poNumber || !poDate || !rateByCustomer || !recordDetails) {
      alert('Please fill all fields.');
      return;
    }

    try {
      const web = new Web("https://skgroupenginering.sharepoint.com/sites/SalesManagement");

      const existingPO = await web.lists
        .getByTitle('POList')
        .items.filter(`SerialNumber eq '${recordDetails.serialNumber}'`)
        .get();

      if (existingPO.length > 0) {
        await web.lists
          .getByTitle('POList')
          .items.getById(existingPO[0].ID)
          .update({
            PONumber: poNumber,
            PODate: poDate,
            RateByCustomer: rateByCustomer,
          });

        alert('PO updated successfully.');
      } else {
        await web.lists
          .getByTitle('POList')
          .items.add({
            PONumber: poNumber,
            PODate: poDate,
            RateByCustomer: rateByCustomer,
            SerialNumber: recordDetails.serialNumber,
          });

        alert('PO created successfully.');
      }

      setShowPOForm(false);
    } catch (error) {
      console.error('Error saving PO:', error);
      alert('Failed to save PO. Please try again.');
    }
  };
  useMemo(async () => {
    if (records.length === 0) return;

    try {
      const web = new Web("https://skgroupenginering.sharepoint.com/sites/SalesManagement");

      const rfqSerialNumbers = records.map((record) => record.serialNumber);
      // console.log("Hooked called of " ,rfqSerialNumbers )
      const poItems = await web.lists
        .getByTitle('POList')
        .items.filter(rfqSerialNumbers.map(num => `SerialNumber eq '${num}'`).join(' or '))
        .expand('AttachmentFiles')
        .select('ID', 'SerialNumber', 'PONumber', 'PODate', 'RateByCustomer', 'AttachmentFiles')
        .get();
        // console.log("List accessed ",poItems)
      const updatedRecords = records.map((record) => {
        const poItem = poItems.find((po) => po.SerialNumber === record.serialNumber);

        if (poItem) {
          return {
            ...record,
            poNumber: poItem.PONumber,
            poDate: poItem.PODate,
            rateByCustomer: poItem.RateByCustomer,
            pdfUrl: poItem.AttachmentFiles.length > 0 ? poItem.AttachmentFiles[0].ServerRelativeUrl : '',
          };
        }

        return record;
      });
      // console.log("Record updated")
      setRecords(updatedRecords);
    } catch (error) {
      console.error('Error fetching PO data for RFQs:', error);
    }
  }, [records]);
  
  
  
  
  
  const handlePdfUpload = async (event: React.ChangeEvent<HTMLInputElement>, serialNumber: string) => {
    if (!event.target.files || event.target.files.length === 0) return;

    const file = event.target.files[0];

    try {
      const web = new Web("https://skgroupenginering.sharepoint.com/sites/SalesManagement");

      const poItems = await web.lists
        .getByTitle('POList')
        .items.filter(`SerialNumber eq '${serialNumber}'`)
        .get();

      if (poItems.length === 0) {
        alert('PO record not found for this Serial Number.');
        return;
      }

      const itemId = poItems[0].ID;

      await web.lists
        .getByTitle('POList')
        .items.getById(itemId)
        .attachmentFiles.add(file.name, file);

      alert('PDF uploaded successfully.');
    } catch (error) {
      console.error('Error uploading PDF:', error);
      alert('Failed to upload PDF. Please try again.');
    }
  };

  return (
    <section className={styles.pOscreen}>
      <h2>Po Management Screen</h2>

      {records.length > 0 && (
        <div className={styles.tableWrapper}>
          <table className={styles.detailsTable}>
            <thead>
              <tr>
                <th>RFQ Number</th>
                {/* <th>Total Rate</th> */}
                <th>Total Weight</th>
                <th>Total Amount</th>
                <th>PO Number</th>
                <th>PO Date</th>
                <th>Final Rate by Customer</th>
                <th>Uploaded PDF</th>
                <th>Action</th>
              </tr>
            </thead>
            <tbody>
  {records.map((record) => (
    <tr key={record.serialNumber}>
      <td>{record.rfqNumber}</td>
      <td>{record.totalWeight}</td>
      <td>{record.totalAmount}</td>
      <td>{record.poNumber || 'N/A'}</td>
      <td>{record.poDate || 'N/A'}</td>
      <td>{record.rateByCustomer || 'N/A'}</td>
      <td>
        {record.pdfUrl ? (
          <a href={record.pdfUrl} target="_blank" rel="noopener noreferrer">
            View PDF
          </a>
        ) : (
          'No PDF Uploaded'
        )}
      </td>
      <td>
        <button className={styles.addButton} 
          onClick={() => {
            setShowPOForm(true);
            setRecordDetails(record);
            setPoNumber(record.poNumber || '');
            setPoDate(record.poDate || '');
            setRateByCustomer(record.rateByCustomer || '');
          }}
        >
          Add PO
        </button>
      </td>
    </tr>
  ))}
</tbody>
          </table>
        </div>
      )}

{showPOForm && (
  <div className={styles.formOverlay}>
    <div className={styles.poForm}>
      <h3>Add Purchase Order</h3>
      <label>
        PO Number:
        <input type="text" name="poNumber" value={poNumber} onChange={handlePOFormChange} />
      </label>
      <label>
        PO Date:
        <input type="date" name="poDate" value={poDate} onChange={handlePOFormChange} />
      </label>
      <label>
        Final Rate by Customer:
        <input type="text" name="rateByCustomer" value={rateByCustomer} onChange={handlePOFormChange} />
      </label>
      <div className={styles.fileUploadWrapper}>
      <label htmlFor="fileUpload" className="customFileLabel">
        Add PO Pdf Given By The Customer
      </label>
      <input
        type="file"
        id="fileUpload"
        accept="application/pdf"
        onChange={(event) => {
          if (recordDetails) {
            handlePdfUpload(event, recordDetails.serialNumber);
            const fileName = event.target.files?.[0]?.name
            document.querySelector('.fileName')!.textContent = fileName;
          }
        }}
      />
    </div>

      <div className={styles.formActions}>
        <button className={styles.saveButton}onClick={savePO}>Save</button>
        <button className={styles.cancelButton}onClick={() => setShowPOForm(false)}>Cancel</button>
      </div>
    </div>
  </div>
)}

    </section>
  );
};

export default POscreen;
