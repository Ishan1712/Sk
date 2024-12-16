import * as React from 'react';
import styles from './Quotationrevisionscreen.module.scss';
import { IQuotationrevisionscreenProps } from './IQuotationrevisionscreenProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { Web }  from 'sp-pnp-js';
import jsPDF from 'jspdf';
import 'jspdf-autotable';

declare module 'jspdf' {
  interface jsPDF {
    autoTable: any;
    lastAutoTable?: { finalY: number };
  }
}

interface IPart {
  partName: string;
  grade: string;
  weight: string;
  overhead: string;
  rate: string;
  labour: string;
  laserCut: string;
  material: string; 
  primer: string;
  quantity: string;
}
interface ICustomerDetails {
  name: string;
  address: string;
  email: string;
  gstNumber: string;
  contactPerson: string;
  mobileNumbers: string[]; // Array of mobile numbers
}

interface IDrawing {
  dlist: string;
  dno: string;
  dquan: string;
  partList: IPart[];
}

interface IQuotationRecord {
  id: number;
  serialNumber: string;
  rfqNumber: string;
  revisionNumber: string;
  quotationDate: string;
  revisionDate: string
  status: string;
  drawingDetails: IDrawing[];
  totalweight: number;
  totalRate: number;
  totalAmount: number;
  customerDetails?: ICustomerDetails;
}

interface IQuotationRevisionScreenState {
  records: IQuotationRecord[];
  isEditing: boolean;
  rfqNumbers: string[];
  selectedSerialNumber: string;
  currentRecord: IQuotationRecord | null;
  showSubmitConfirm: boolean;
  showDeleteConfirm: boolean;
  recordToConfirm: IQuotationRecord | null;
  selectedDrawingIndex: number | null;
}

export default class Quotationrevisionscreen extends React.Component<IQuotationrevisionscreenProps, IQuotationRevisionScreenState> {
  constructor(props: IQuotationrevisionscreenProps) {
    super(props);
    this.state = {
      records: [],
      isEditing: false,
      rfqNumbers: [],
      selectedSerialNumber: "",
      currentRecord: null,
      showSubmitConfirm: false,
      showDeleteConfirm: false,
      recordToConfirm: null,
      selectedDrawingIndex: 0,
    };
  }

  private loadRFQNumbersFromSharePoint = async (): Promise<void> => {
    try {
      const web = new Web("https://skgroupenginering.sharepoint.com/sites/SalesManagement"); 
      const items = await web.lists
        .getByTitle("RFQList")
        .items.select("RFQNumber")
        .filter("Status eq 'Revised'")
        .get();
  
      const rfqNumbers = items.map((item: any) => item.RFQNumber);
      this.setState({ rfqNumbers });
    } catch (error) {
      console.error("Error loading RFQ numbers:", error);
    }
  };
  private fetchDrawingAndPartDetailsBySerialNumber = async (rfqNumber: string): Promise<IDrawing[]> => {
    try {
      const web = new Web("https://skgroupenginering.sharepoint.com/sites/SalesManagement");
  
      // Fetch drawings associated with the RFQ number
      const drawingItems = await web.lists
        .getByTitle("DrawingList")
        .items.filter(`RFQNumber eq '${rfqNumber}'`)
        .select("DrawingNumber", "DrawingQuantity")
        .get();
  
      // Fetch parts associated with the RFQ number
      const partItems = await web.lists
        .getByTitle("PartList")
        .items.filter(`RFQNumber eq '${rfqNumber}'`)
        .select(
          "PartName",
          "Material",
          "Grade",
          "Weight",
          "Overhead",
          "Rate",
          "Labour",
          "LaserCut",
          "Primer",
          "DrawingNumber" // Ensure this column is present in PartList
        )
        .get();
  
      // Map drawing details and include part details
      const drawingDetails: IDrawing[] = drawingItems.map((drawing: any) => ({
        dlist: `Drawing ${drawing.DrawingNumber}`,
        dno: drawing.DrawingNumber,
        dquan: drawing.DrawingQuantity,
        partList: partItems.filter((part: any) => part.DrawingNumber === drawing.DrawingNumber).map((part: any) => ({
          partName: part.PartName,
          material: part.Material || "",
          grade: part.Grade || "",
          weight: part.Weight || "",
          overhead: part.Overhead || "",
          rate: part.Rate || "",
          labour: part.Labour || "",
          laserCut: part.LaserCut || "",
          primer: part.Primer || "",
        })),
      }));
  
      console.log("Drawing and Part Details Fetched:", drawingDetails);
      return drawingDetails;
    } catch (error) {
      console.error("Error fetching drawing and part details:", error);
      return [];
    }
  };
  
  private fetchQuotationDetails = async (rfqNumber: string): Promise<{ 
    quotationDate: string; 
    totalWeight: number; 
    totalRate: number; 
    totalAmount: number; 
    status: string; 
  }> => {
    try {
      const web = new Web("https://skgroupenginering.sharepoint.com/sites/SalesManagement"); 
  
      const quotationItems = await web.lists
        .getByTitle("QuotationList")
        .items.filter(`RFQSerialNumber eq '${rfqNumber}'`)
        .select("QuotationDate", "TotalWeight", "TotalRate", "TotalAmount", "Status")
        .get();
  
      if (quotationItems.length > 0) {
        const item = quotationItems[0];
        return {
          quotationDate: item.QuotationDate || "",
          totalWeight: parseFloat(item.TotalWeight || "0"),
          totalRate: parseFloat(item.TotalRate || "0"),
          totalAmount: parseFloat(item.TotalAmount || "0"),
          status: item.Status || "Pending",
        };
      }
  
      return { quotationDate: "", totalWeight: 0, totalRate: 0, totalAmount: 0, status: "Pending" };
    } catch (error) {
      console.error("Error fetching quotation details:", error);
      return { quotationDate: "", totalWeight: 0, totalRate: 0, totalAmount: 0, status: "Pending" };
    }
  };
  
  
  
  private fetchCustomerDetailsForSpecificRFQ = async (rfqNumber: string): Promise<ICustomerDetails | null> => {
    try {
      const web = new Web("https://skgroupenginering.sharepoint.com/sites/SalesManagement"); 
  
      // Fetch RFQ details to retrieve the associated customer name
      const rfqItem = await web.lists
        .getByTitle("RFQList")
        .items.filter(`RFQNumber eq '${rfqNumber}'`)
        .select("CustomerName")
        .get();
  
      if (!rfqItem || rfqItem.length === 0) {
        console.warn(`No RFQ data found for RFQNumber: ${rfqNumber}`);
        return null;
      }
  
      const customerName = rfqItem[0].CustomerName;
  
      // Fetch customer details for the retrieved customer name
      const customerItem = await web.lists
        .getByTitle("CustomerList")
        .items.filter(`CustomerName eq '${customerName}'`)
        .select("CustomerName", "Address", "Email", "GSTNumber", "ContactPerson", "MobileNumber")
        .get();
  
      if (!customerItem || customerItem.length === 0) {
        console.warn(`No customer data found for CustomerName: ${customerName}`);
        return null;
      }
  
      const customer = customerItem[0];
      return {
        name: customer.CustomerName,
        address: customer.Address,
        email: customer.Email,
        gstNumber: customer.GSTNumber,
        contactPerson: customer.ContactPerson,
        mobileNumbers: customer.MobileNumber.split(", "),
      };
    } catch (error) {
      console.error("Error fetching customer details for specific RFQ:", error);
      return null;
    }
  };

  private updatePartDetailsInList = async (
    rfqNumber: string,
    partName: string,
    updatedPart: Partial<IPart>
  ): Promise<void> => {
    try {
      const web = new Web("https://skgroupenginering.sharepoint.com/sites/SalesManagement");
  
      // Get the specific part item based on RFQNumber and PartName
      const items = await web.lists
        .getByTitle("PartList")
        .items.filter(`RFQNumber eq '${rfqNumber}' and PartName eq '${partName}'`)
        .select("Id") // Fetch only the ID to perform the update
        .get();
  
      if (items.length === 0) {
        console.warn(`No matching part found for RFQNumber: ${rfqNumber} and PartName: ${partName}`);
        return;
      }
      const itemId = items[0].Id; // Assuming PartName is unique within the RFQNumber context
  
      // Update the part details
      await web.lists.getByTitle("PartList").items.getById(itemId).update({
        Grade: updatedPart.grade,
        Weight: updatedPart.weight,
        Overhead: updatedPart.overhead,
        Rate: updatedPart.rate,
        Labour: updatedPart.labour,
        LaserCut: updatedPart.laserCut,
        Primer: updatedPart.primer,
      });
  
      alert(`Part details for "${partName}" updated successfully!`);
    } catch (error) {
      console.error("Error updating part details in the list:", error);
      alert("Failed to update part details. Please try again.");
    }
  };


  private handleDownloadPDF = (serialNumber: string) => {
    try {
      // Find the index of the record using a loop
      let selectedIndex = -1;
      for (let i = 0; i < this.state.records.length; i++) {
        if (this.state.records[i].serialNumber === serialNumber) {
          selectedIndex = i;
          break;
        }
      }
  
      const selectedRecord = selectedIndex !== -1 ? this.state.records[selectedIndex] : null;
  
      if (!selectedRecord) {
        alert("No record found for the selected RFQ number.");
        return;
      }
  
      const doc = new jsPDF('landscape');
      const pageWidth = doc.internal.pageSize.getWidth();
  
      // Add Header
      doc.setFontSize(12);
      doc.text('S.K. GROUP ENGINEERING', pageWidth / 2, 10, { align: 'center' });
      doc.setFontSize(10);
      doc.text('Gat No. 240, Dhanore, Vikaswadi, Near Dhanore Phata, Markal Road, Tal Khed, Distt. Pune', pageWidth / 2, 16, { align: 'center' });
      doc.text('Pin No. 412105', pageWidth / 2, 20, { align: 'center' });
      doc.text('E-mail: enquiry@skgroupengineering.com | Cell No: 9960414239', pageWidth / 2, 24, { align: 'center' });
  
      // Table Headers with new columns
      const headers = [
        [
          'SR NO',
          'DRG NO.',
          'ITEM',
          'QTY',
          'GRADE',
          'WT',
          'OH',
          'T.WT',
          'RATE',
          'LABOUR',
          'L/C',
          'PRIMER',
          'T.RATE',
          'AMOUNT',
          'REVISION NO',
          'QUOTATION DATE',
          'REVISION DATE',
        ]
      ];
  
      const rows: any[] = [];
      let srNo = 1;
  
      // Populate Table Rows
      selectedRecord.drawingDetails.forEach((drawing) => {
        rows.push([
          `${srNo}`,
          drawing.dno,
          '', // ITEM is blank for the drawing row
          '', '', '', '', '', '', '', '', '', '', '',
          selectedRecord.revisionNumber || '-', // Add Revision Number
          selectedRecord.quotationDate ? new Date(selectedRecord.quotationDate).toISOString().split('T')[0] : '-', // Add Quotation Date
          selectedRecord.revisionDate || '-', // Add Revision Date
        ]);
  
        drawing.partList.forEach((part) => {
          rows.push([
            '', // Blank SR NO for parts
            '', // Blank DRG NO for parts
            part.partName,
            part.quantity,
            part.grade,
            part.weight,
            part.overhead,
            this.totalWeight(part).toFixed(2),
            part.rate,
            part.labour,
            part.laserCut,
            part.primer,
            this.totalRate(part).toFixed(2),
            this.calculatePartTotal(part).toFixed(2),
            '', // Blank for revision details in part rows
            '', // Blank for quotation date
            '', // Blank for revision date
          ]);
        });
  
        srNo++;
      });
  
      // Add Totals Row
      rows.push([
        '', // Blank SR NO
        '', // Blank DRG NO
        'TOTAL',
        '', '', '', '',
        selectedRecord.totalweight.toFixed(2),
        '', '', '',
        '', // Primer total is not calculated; leave blank or adjust if needed
        selectedRecord.totalRate.toFixed(2),
        selectedRecord.totalAmount.toFixed(2),
        '', // Blank for revision number
        '', // Blank for quotation date
        '', // Blank for revision date
      ]);
  
      // Adjust Column Widths
      const columnWidths = [
        10, // SR NO
        15, // DRG NO
        30, // ITEM
        15, // QTY
        17, // GRADE
        15, // WT
        15, // OH
        15, // T.WT
        13, // RATE
        13, // LABOUR
        13, // L/C
        13, // PRIMER
        18, // T.RATE
        20, // AMOUNT
        15, // REVISION NO
        23, // QUOTATION DATE
        23, // REVISION DATE
      ];
  
      // Generate Table
      doc.autoTable({
        head: headers,
        body: rows,
        startY: 30,
        columnStyles: columnWidths.reduce((acc, width, index) => ({ ...acc, [index]: { cellWidth: width } }), {}),
        styles: {
          fontSize: 8,
          cellPadding: 2,
        },
        headStyles: {
          fillColor: [0, 102, 204], // Blue header
          textColor: 255,
          halign: 'center',
        },
        bodyStyles: {
          halign: 'center',
          valign: 'middle',
        },
      });
  
      // Save PDF
      doc.save(`${serialNumber}_revision_report.pdf`);
    } catch (error) {
      console.error("Error generating PDF:", error);
    }
  };

  private handleDownloadSecondPDF = async (serialNumber: string) => {
    try {
      const web = new Web("https://skgroupenginering.sharepoint.com/sites/SalesManagement");
  
      // Fetch RFQ details from SharePoint
      const rfqItems = await web.lists
        .getByTitle("RFQList")
        .items.filter(`RFQNumber eq '${serialNumber}'`)
        .select("RFQNumber", "CustomerName", "Date", "Subject")
        .get();
  
      if (!rfqItems || rfqItems.length === 0) {
        alert("No RFQ data found for the selected RFQ number.");
        return;
      }
  
      const rfq = rfqItems[0];
      const customerName = rfq.CustomerName ?? "N/A";
  
      // Fetch ContactPerson from the CustomerList
      const customerItem = await web.lists
        .getByTitle("CustomerList")
        .items.filter(`CustomerName eq '${customerName}'`)
        .select("ContactPerson")
        .get();
  
      const contactPerson = customerItem.length > 0 ? customerItem[0].ContactPerson : "N/A";
  
      // Format other details
      const formattedDate = rfq.Date
        ? new Date(rfq.Date).toLocaleDateString("en-GB") // Format to DD-MM-YYYY
        : "N/A";
  
      const refNumber = `SKG/${rfq.RFQNumber}`;
      const subject = rfq.Subject ?? "N/A";
  
      const doc = new jsPDF();
  
      // Add Header
      doc.setFontSize(12);
      doc.text("S.K. GROUP ENGINEERING", 105, 10, { align: "center" });
      doc.setFontSize(10);
      doc.text(
        "Gate no: 240, DanoreVikaswadi, Near Dhanore Phata, Markal Road Tal- Khed, Dist Pune - 412105",
        105,
        16,
        { align: "center" }
      );
      doc.text("Email: skgupta.07sg@gmail.com | Contact: 9960414239", 105, 22, {
        align: "center",
      });
  
      // Add RFQ Details
      doc.setFontSize(11);
      doc.text(`Ref: ${refNumber}`, 14, 40);
      doc.text(`Date: ${formattedDate}`, 160, 40);
      doc.text(`TO:`, 14, 50);
      doc.text(`${customerName}`, 14, 56);
  
      doc.text(`SUBJECT: ${subject}`, 14, 66);
      doc.text(`KIND ATTN: ${contactPerson}`, 14, 74);
  
      doc.setFontSize(11);
      doc.text(
        `DEAR SIR, WITH REFERENCE TO YOUR ENQUIRY RECEIVED & SUBSEQUENT DISCUSSIONS HAD WITH YOU, WE ARE PLEASED TO SUBMIT OUR BUDGETARY OFFER AS FOLLOWS:`,
        14,
        84,
        { maxWidth: 180 }
      );
  
      // Add Scope of Work Section
      doc.setFontSize(12);
      doc.text("SCOPE OF WORK:", 14, 100);
  
      doc.setFontSize(10);
      doc.text(`WE OFFER OUR BEST QUOTATION FOR ${subject.toUpperCase()}.`, 14, 110);
  
      // Use indexOf to find the correct record
      let selectedIndex = -1;
      for (let i = 0; i < this.state.records.length; i++) {
        if (this.state.records[i].serialNumber === serialNumber) {
          selectedIndex = i;
          break;
        }
      }
  
      if (selectedIndex === -1) {
        alert("No record found for the selected RFQ number.");
        return;
      }
  
      const selectedRecord = this.state.records[selectedIndex];
  
      // Add Table
      const tableHeaders = [["SR NO", "DRG NO.", "T.WT", "OH", "RATE", "AMOUNT"]];
      const tableRows = selectedRecord.drawingDetails.map((drawing, index) => {
        const { totalWeight, totalRate, totalAmount } = this.calculateDrawingTotals(drawing);
        const overhead = totalWeight * 0.1; // Overhead as 10% of total weight
  
        return [
          `${index + 1}`,
          drawing.dno,
          totalWeight.toFixed(2),
          overhead.toFixed(2),
          totalRate.toFixed(2),
          totalAmount.toFixed(2),
        ];
      });
  
      const startY = 120;
      doc.autoTable({
        head: tableHeaders,
        body: tableRows,
        startY,
      });
  
      // Add Footer Terms
      const finalY = doc.lastAutoTable?.finalY ?? startY + 10;
      doc.setFontSize(11);
      doc.text("TERMS & CONDITIONS:", 14, finalY + 10);
      doc.text(
        "The prices are works Alandi, Pune basis.\nTAXES: GST 18% Extra at actual.\nDELIVERY: Within 3-4 weeks from the date of receipt of your order.\nTRANSPORT: Extra at actual (to be paid by customer).\nVALIDITY: 10 days.\nPAYMENT: 50% Advance, Balance 50% + GST against dispatch.",
        14,
        finalY + 20,
        { maxWidth: 180 }
      );
  
      // Add Closing Statement
      doc.setFontSize(11);
      doc.text("We hope that the above is as per your requirement.", 14, finalY + 90);
      doc.text("Awaiting for your valued purchase order.", 14, finalY + 96);
  
      doc.setFontSize(12);
      doc.text("With kind regards,", 14, finalY + 110);
      doc.text("Sanjay Gupta", 14, finalY + 116);
      doc.text("9960414239", 14, finalY + 122);
  
      // Save PDF
      doc.save(`${serialNumber}_Quotation.pdf`);
    } catch (error) {
      console.error("Error generating PDF:", error);
      alert("Failed to generate PDF. Please try again.");
    }
  };


  private handleDownloadBothPDFs = async (serialNumber: string) => {
    try {
      // Call the first PDF generation
      await this.handleDownloadPDF(serialNumber);
  
      // Call the second PDF generation
      await this.handleDownloadSecondPDF(serialNumber);
      
      alert("Both PDFs downloaded successfully.");
    } catch (error) {
      console.error("Error downloading PDFs:", error);
      alert("Failed to download PDFs. Please try again.");
    }
  };
    
  public componentDidMount(): void {
    this.loadRFQNumbersFromSharePoint();
  }

  private handleSerialNumberChange = async (e: React.ChangeEvent<HTMLSelectElement>) => {
    const serialNumber = e.target.value;
  
    if (!serialNumber) return;
  
    try {
      const drawingDetails = await this.fetchDrawingAndPartDetailsBySerialNumber(serialNumber);
      const { quotationDate,totalWeight, totalRate, totalAmount, status } = await this.fetchQuotationDetails(serialNumber);
      const customerDetails = await this.fetchCustomerDetailsForSpecificRFQ(serialNumber);
  
      const currentDate = new Date().toISOString().split("T")[0];
  
      let existingRecordIndex = -1;
      for (let i = 0; i < this.state.records.length; i++) {
        if (this.state.records[i].serialNumber === serialNumber) {
          existingRecordIndex = i;
          break;
        }
      }
  
      let updatedRecords = [...this.state.records];
  
      if (existingRecordIndex >= 0) {
        // Update existing record
        updatedRecords[existingRecordIndex] = {
          ...updatedRecords[existingRecordIndex],
          drawingDetails,
          customerDetails: customerDetails || updatedRecords[existingRecordIndex].customerDetails,
          revisionNumber: (
            parseInt(updatedRecords[existingRecordIndex].revisionNumber || "0") + 1
          ).toString(),
          quotationDate,
          status,
          totalweight: totalWeight,
          totalRate: totalRate,
          totalAmount: totalAmount,
          revisionDate: currentDate,
        };
      } else {
        // Add new record
        const newRecord: IQuotationRecord = {
          id: updatedRecords.length + 1,
          serialNumber,
          rfqNumber: serialNumber,
          revisionNumber: "1",
          quotationDate,
          revisionDate: currentDate,
          status,
          drawingDetails,
          totalweight: totalWeight,
          totalRate: totalRate,
          totalAmount: totalAmount,
          customerDetails: customerDetails || undefined,
        };
        updatedRecords.push(newRecord);
      }
  
      this.setState({
        selectedSerialNumber: serialNumber,
        records: updatedRecords,
      });
    } catch (error) {
      console.error("Error handling serial number change:", error);
    }
  };
  
  
  private handleEdit = (record: IQuotationRecord) => {
    this.setState({ isEditing: true, currentRecord: { ...record }, selectedDrawingIndex: 0 });
  };

  private handleDrawingChange = (e: React.ChangeEvent<HTMLSelectElement>) => {
    this.setState({ selectedDrawingIndex: parseInt(e.target.value) });
  };

  private handleSubmitConfirm = (record: IQuotationRecord) => {
    this.setState({ showSubmitConfirm: true, recordToConfirm: record });
  };

  private confirmSubmit = async (): Promise<void> => {
    const { recordToConfirm, records } = this.state;
  
    if (recordToConfirm) {
      try {
        const web = new Web("https://skgroupenginering.sharepoint.com/sites/SalesManagement") // Add your SharePoint URL here
  
        // Fetch the quotation record by RFQNumber
        const quotationItems = await web.lists
          .getByTitle("QuotationList")
          .items.filter(`RFQSerialNumber eq '${recordToConfirm.rfqNumber}'`)
          .select("Id")
          .get();
  
        if (quotationItems.length > 0) {
          const quotationItemId = quotationItems[0].Id;
  
          // Update the status to "WorkingDone" in the QuotationList
          await web.lists.getByTitle("QuotationList").items.getById(quotationItemId).update({
            Status: "WorkingDone",
          });
  
          // Fetch the RFQ item ID from the RFQList using the RFQNumber
          const rfqItems = await web.lists
            .getByTitle("RFQList")
            .items.filter(`RFQNumber eq '${recordToConfirm.rfqNumber}'`)
            .select("Id")
            .get();
  
          if (rfqItems.length > 0) {
            const rfqItemId = rfqItems[0].Id;
  
            // Update the status to "WorkingDone" in the RFQList
            await web.lists.getByTitle("RFQList").items.getById(rfqItemId).update({
              Status: "WorkingDone",
            });
          }
  
          // Update local state
          const updatedRecords = records.map((rec) =>
            rec.id === recordToConfirm.id ? { ...rec, status: "WorkingDone" } : rec
          );
  
          this.setState({
            records: updatedRecords,
            showSubmitConfirm: false,
            recordToConfirm: null,
          });
  
          alert(`Quotation status updated to "WorkingDone" successfully.`);
        } else {
          alert("Quotation record not found in the SharePoint list.");
        }
      } catch (error) {
        console.error("Error updating quotation status:", error);
        alert("Failed to update the quotation status. Please try again.");
      }
    }
  };
  
  

  private cancelSubmit = () => {
    this.setState({ showSubmitConfirm: false, recordToConfirm: null });
  };

  private handleDeleteConfirm = (record: IQuotationRecord) => {
    this.setState({ showDeleteConfirm: true, recordToConfirm: record });
  };

  private confirmDelete = () => {
    const { recordToConfirm, records } = this.state;
    if (recordToConfirm) {
      const updatedRecords = records.filter(record => record.id !== recordToConfirm.id);
      this.setState({ records: updatedRecords, showDeleteConfirm: false, recordToConfirm: null });
    }
  };

  private cancelDelete = () => {
    this.setState({ showDeleteConfirm: false, recordToConfirm: null });
  };

  private saveEdit = async (): Promise<void> => {
    const { currentRecord, records } = this.state;
  
    if (currentRecord) {
      const rfqNumber = currentRecord.rfqNumber;
  
      try {
        const web = new Web("https://skgroupenginering.sharepoint.com/sites/SalesManagement"); 
  
        // Update QuotationList
        const quotationItems = await web.lists
          .getByTitle("QuotationList")
          .items.filter(`RFQSerialNumber eq '${rfqNumber}'`)
          .select("Id")
          .get();
  
        if (quotationItems.length > 0) {
          const itemId = quotationItems[0].Id;
          await web.lists.getByTitle("QuotationList").items.getById(itemId).update({
            QuotationDate: currentRecord.quotationDate,
            RevisionNumber: currentRecord.revisionNumber,
          });
        }
  
        // Update PartList using updatePartDetailsInList
        for (const drawing of currentRecord.drawingDetails) {
          for (const part of drawing.partList) {
            await this.updatePartDetailsInList(rfqNumber, part.partName, {
              grade: part.grade,
              weight: part.weight,
              overhead: part.overhead,
              rate: part.rate,
              labour: part.labour,
              laserCut: part.laserCut,
              primer: part.primer,
            });
          }
        }
  
        // Update local state
        const updatedRecords = records.map((record) =>
          record.id === currentRecord.id ? currentRecord : record
        );
        this.setState({ records: updatedRecords, isEditing: false, currentRecord: null });
        alert("Changes saved successfully!");
      } catch (error) {
        console.error("Error saving edits:", error.message || error);
        alert(`Failed to save changes: ${error.message || error}`);
      }
    }
  };
  
  
  private cancelEdit = () => {
    this.setState({ isEditing: false, currentRecord: null });
  };

  private handleChange = (
    e: React.ChangeEvent<HTMLInputElement>,
    drawingIndex?: number,
    partIndex?: number
  ) => {
    const { name, value } = e.target;
    this.setState((prevState) => {
      if (drawingIndex !== undefined && partIndex !== undefined) {
        const updatedRecord = { ...prevState.currentRecord };
        updatedRecord!.drawingDetails[drawingIndex].partList[partIndex] = {
          ...updatedRecord!.drawingDetails[drawingIndex].partList[partIndex],
          [name]: value,
        };
        return { currentRecord: updatedRecord };
      }
      return {
        currentRecord: {
          ...prevState.currentRecord!,
          [name]: value,
        },
      };
    });
  };
  
  private calculateDrawingTotals = (drawing: IDrawing): { totalWeight: number; totalRate: number; totalAmount: number } => {
    const totalWeight = drawing.partList.reduce((sum, part) => sum + this.totalWeight(part), 0);
    const totalRate = drawing.partList.reduce((sum, part) => sum + this.totalRate(part), 0);
    const totalAmount = drawing.partList.reduce((sum, part) => sum + this.calculatePartTotal(part), 0);
    return { totalWeight, totalRate, totalAmount };
  };

  private calculatePartTotal = (part: IPart): number => {
    const weight = parseFloat(part.weight || '0');
    const overhead = parseFloat(part.overhead || '0');
    const rate = parseFloat(part.rate || '0');
    const labour = parseFloat(part.labour || '0');
    const laserCut = parseFloat(part.laserCut || '0');
    const primer = parseFloat(part.primer || '0');

    const totalRate = rate + labour + laserCut + primer;
    const totalWeight = weight + overhead;
    return totalRate * totalWeight;
  };

  private totalWeight = (part: IPart): number => {
    const weight = parseFloat(part.weight || '0');
    const overhead = parseFloat(part.overhead || '0');

    return weight + overhead;
  };

  private totalRate = (part: IPart): number => {
    const rate = parseFloat(part.rate || '0');
    const labour = parseFloat(part.labour || '0');
    const laserCut = parseFloat(part.laserCut || '0');
    const primer = parseFloat(part.primer || '0');

    return rate + labour + laserCut + primer;
  };

  private calculateTotalWeight = (record: IQuotationRecord): number => {
    return record.drawingDetails.reduce((totalWeight, drawing) => {
      return totalWeight + drawing.partList.reduce((partWeight, part) => {
        return partWeight + this.totalWeight(part);
      }, 0);
    }, 0);
  };

  private calculateTotalRate = (record: IQuotationRecord): number => {
    return record.drawingDetails.reduce((totalRate, drawing) => {
      return totalRate + drawing.partList.reduce((partRate, part) => {
        return partRate + this.totalRate(part);
      }, 0);
    }, 0);
  };

  private calculateTotalAmount = (record: IQuotationRecord): number => {
    return record.drawingDetails.reduce((totalAmount, drawing) => {
      return totalAmount + drawing.partList.reduce((partAmount, part) => {
        return partAmount + this.calculatePartTotal(part);
      }, 0);
    }, 0);
  };

  public render(): React.ReactElement<IQuotationrevisionscreenProps> {
    const { records, isEditing, currentRecord, selectedSerialNumber, showSubmitConfirm, showDeleteConfirm } = this.state;

    return (
      <section className={styles.quotationrevisionscreen}>
        <h2>Quotation Revision Screen</h2>
        <div className={styles.rfqBox}>
  <label htmlFor="rfqSelect">Select RFQ Serial Number:</label>
  <select
    id="rfqSelect"
    value={this.state.selectedSerialNumber}
    onChange={this.handleSerialNumberChange}
  >
    <option value="">Select...</option>
    {this.state.rfqNumbers.map((rfqNumber, index) => (
      <option key={index} value={rfqNumber}>
        {rfqNumber}
      </option>
    ))}
  </select>
</div>




        {showSubmitConfirm && (
          <div className={styles.confirmOverlay}>
            <div className={styles.confirmBox}>
              <p>Are you sure you haved Revised this data?</p>
              <button onClick={this.confirmSubmit} className={styles.confirmButton}>Yes</button>
              <button onClick={this.cancelSubmit} className={styles.cancelButton}>No</button>
            </div>
          </div>
        )}

        {showDeleteConfirm && (
          <div className={styles.confirmOverlay}>
            <div className={styles.confirmBox}>
              <p>Are you sure you want to delete this data?</p>
              <button onClick={this.confirmDelete} className={styles.confirmButton}>Yes</button>
              <button onClick={this.cancelDelete} className={styles.cancelButton}>No</button>
            </div>
          </div>
        )}

        {isEditing && currentRecord && (
          <div className={styles.editOverlay}>
            <div className={styles.editBox}>
              <h3>Edit Record</h3>
              <div>
                <label>
                  Quotation Date:
                  <input type="date" name="quotationDate" value={currentRecord.quotationDate} onChange={this.handleChange} />
                </label>
                <label>
                <label>
                  Revision Date:
                  <input type="date" name="revisionDate" value={currentRecord.revisionDate} onChange={this.handleChange} />
                </label>
                <label></label>
                  Revision Number:
                  <input type="text" name="revisionNumber" value={currentRecord.revisionNumber} onChange={this.handleChange} />
                </label>
                <label>
                  Select Drawing:
                  <select className={styles.drawingDropdown} value={this.state.selectedDrawingIndex ?? ""} onChange={this.handleDrawingChange}>
                    {currentRecord.drawingDetails.map((drawing, index) => (
                      <option key={index} value={index}>{drawing.dlist}</option>
                    ))}
                  </select>
                </label>
                {this.state.selectedDrawingIndex !== null &&
                  currentRecord.drawingDetails[this.state.selectedDrawingIndex].partList.map((part, pIndex) => (
                    <div key={pIndex} className={styles.partEdit}>
                      <label>Part Name: <input type="text" value={part.partName} readOnly /></label>
                      <label>Grade: <input type="text" name="grade" value={part.grade} onChange={(e) => this.handleChange(e, this.state.selectedDrawingIndex!, pIndex)} /></label>
                      <label>Weight: <input type="text" name="weight" value={part.weight} onChange={(e) => this.handleChange(e, this.state.selectedDrawingIndex!, pIndex)} /></label>
                      <label>Overhead: <input type="text" name="overhead" value={part.overhead} onChange={(e) => this.handleChange(e, this.state.selectedDrawingIndex!, pIndex)} /></label>
                      <label>Rate: <input type="text" name="rate" value={part.rate} onChange={(e) => this.handleChange(e, this.state.selectedDrawingIndex!, pIndex)} /></label>
                      <label>Labour: <input type="text" name="labour" value={part.labour} onChange={(e) => this.handleChange(e, this.state.selectedDrawingIndex!, pIndex)} /></label>
                      <label>Laser Cut: <input type="text" name="laserCut" value={part.laserCut} onChange={(e) => this.handleChange(e, this.state.selectedDrawingIndex!, pIndex)} /></label>
                      <label>Primer: <input type="text" name="primer" value={part.primer} onChange={(e) => this.handleChange(e, this.state.selectedDrawingIndex!, pIndex)} /></label>
                    </div>
                  ))}
              </div>
              <button onClick={this.saveEdit} className={styles.saveButton}>Save</button>
              <button onClick={this.cancelEdit} className={styles.cancelButton}>Cancel</button>
            </div>
          </div>
        )}

        <div className={styles.tableWrapper}>
          {records.length > 0 ? (
            <table className={styles.quotationTable}>
            <thead>
              <tr>
                <th>Serial No</th>
                <th>Customer Details</th>
                <th>Drawing & Part Details</th>
                <th>Quotation Date</th> 
                <th>Revision Date</th> 
                <th>Revision Number</th>
                <th>Total Weight</th> 
                <th>Total Rate</th> 
                <th>Total Amount</th> 
                <th>Status</th> 
                <th>Actions</th> 
              </tr>
            </thead>
            <tbody>
        {records.map(record => (
          <tr key={record.id}>
            <td>{record.serialNumber}</td>
            <td>
              {record.customerDetails ? (
                <>
                  <strong>Name:</strong> {record.customerDetails.name}<br />
                  <strong>Address:</strong> {record.customerDetails.address}<br />
                  <strong>Email:</strong> {record.customerDetails.email}<br />
                  <strong>GST:</strong> {record.customerDetails.gstNumber}<br />
                  <strong>Contact:</strong> {record.customerDetails.contactPerson}<br />
                  <strong>Mobile:</strong> {record.customerDetails.mobileNumbers.join(", ")}<br />
                </>
              ) : (
                "No customer details available"
              )}
            </td>
            <td>
        {record.drawingDetails.map((drawing, dIndex) => (
          <div key={dIndex}>
            <strong>{drawing.dlist}</strong> (Dno: {drawing.dno}, Quantity: {drawing.dquan})
            <ul>
              {drawing.partList.map((part, pIndex) => (
                <li key={pIndex}>
                  <strong>{part.partName}</strong> - 
                  Grade: {part.grade}, 
                  Weight: {part.weight}, 
                  Overhead: {part.overhead}, 
                  Rate: {part.rate}, 
                  Labour: {part.labour}, 
                  Laser cut: {part.laserCut}, 
                  Primer: {part.primer}
                  <br />
                  <strong>Total Weight:</strong> {this.totalWeight(part).toFixed(2)}
                  <br />
                  <strong>Total Rate:</strong> {this.totalRate(part).toFixed(2)}
                  <br />
                  <strong>Part Total:</strong> {this.calculatePartTotal(part).toFixed(2)}
                </li>
              ))}
            </ul>
          </div>
        ))}
      </td>
            <td>{record.quotationDate}</td>
            <td>{record.revisionDate}</td>
            <td>{record.revisionNumber}</td>
            <td>{record.totalweight}</td>
            <td>{record.totalRate}</td>
            <td>{record.totalAmount}</td>
            <td>{record.status}</td>
            <td>
            <button onClick={() => this.handleDownloadBothPDFs(record.serialNumber)} className={styles.downloadButton}>Download PDF</button>
            <button
              onClick={() => this.handleEdit(record)}
              className={styles.editButton}
            >
              Edit
            </button>
            <button
              onClick={() => this.handleSubmitConfirm(record)}
              className={styles.submitButton}
            >
              Revised
            </button>
            <button
              onClick={() => this.handleDeleteConfirm(record)}
              className={styles.deleteButton}
            >
              Delete
            </button>
          </td>
          </tr>
        ))}
      </tbody>
      </table>
  ) : (
    <p>No records to display. Select a serial number to add a new row.</p>
  )}
</div>
      </section>
    );
  }
}