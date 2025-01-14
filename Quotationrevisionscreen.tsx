import * as React from 'react';
import styles from './Quotationrevisionscreen.module.scss';
import { IQuotationrevisionscreenProps } from './IQuotationrevisionscreenProps';
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
  totalWeight: number; // Sum of all parts' totalWeight
  avgRate: number; // Average of all parts' totalRate
  totalAmount: number; // Sum of all parts' totalAmount
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
      const web = new Web("https://skgroupenginering.sharepoint.com/sites/SalesManagement")
  
      // Fetch drawings associated with the RFQ number
      const drawingItems = await web.lists
        .getByTitle("DrawingList")
        .items.filter(`RFQNumber eq '${rfqNumber}'`)
        .select("DrawingNumber", "DrawingQuantity", "TotalWeight", "AvgRate", "TotalAmount")
        .get();
  
      // Fetch parts associated with the RFQ number
      const partItems = await web.lists
        .getByTitle("PartList")
        .items.filter(`RFQNumber eq '${rfqNumber}'`)
        .select(
          "PartName",
          "Material",
          "Weight",
          "Overhead",
          "Rate",
          "Labour",
          "LaserCut",
          "Primer",
          "Quantity",
          "DrawingNumber"
        )
        .get();
  
      // Map drawing details and include part details
      const drawingDetails: IDrawing[] = drawingItems.map((drawing: any) => ({
        dlist: `Drawing ${drawing.DrawingNumber}`,
        dno: drawing.DrawingNumber,
        dquan: drawing.DrawingQuantity,
        totalWeight: parseFloat(drawing.TotalWeight || "0"),
        avgRate: parseFloat(drawing.AvgRate || "0"),
        totalAmount: parseFloat(drawing.TotalAmount || "0"),
        partList: partItems
          .filter((part: any) => part.DrawingNumber === drawing.DrawingNumber)
          .map((part: any) => ({
            partName: part.PartName,
            material: part.Material || "",
            weight: part.Weight || "",
            overhead: part.Overhead || "",
            rate: part.Rate || "",
            labour: part.Labour || "",
            laserCut: part.LaserCut || "",
            primer: part.Primer || "",
            quantity: part.Quantity || "0",
          })),
      }));
  
      console.log("Drawing and Part Details Fetched:", drawingDetails);
      return drawingDetails;
    } catch (error) {
      console.error("Error fetching drawing and part details:", error);
      return [];
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
        Material: updatedPart.material,
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


  private handleDownloadPDF = (rfqNumber: string) => {
    const { records } = this.state;
    const selectedRecord = records.filter(record => record.serialNumber === rfqNumber)[0] || null;

    if (!selectedRecord) {
        alert("No record found for the selected RFQ number.");
        return;
    }

    const doc = new jsPDF('landscape'); // Set orientation to landscape
    const pageWidth = doc.internal.pageSize.getWidth();

    // Add Header
    doc.setFontSize(12);
    doc.text('S.K. GROUP ENGINEERING', pageWidth / 2, 10, { align: 'center' });
    doc.setFontSize(10);
    doc.text('Gat No. 240, Dhanore, Vikaswadi, Near Dhanore Phata, Markal Road, Tal Khed, Distt. Pune', pageWidth / 2, 16, { align: 'center' });
    doc.text('Pin No. 412105', pageWidth / 2, 20, { align: 'center' });
    doc.text('E-mail: sonalienterprises89@rediffmail.com | Cell No: 9960414239', pageWidth / 2, 24, { align: 'center' });

    // Table Headers
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
        ]
    ];

    const rows: any[] = [];
    let srNo = 1;
    let totalOverhead = 0;
    let totalWeightSum = 0; // Total of all part weights
    let totalTWeightSum = 0; // Total of all total weights (weight + overhead)

    // Populate Table Rows
    selectedRecord.drawingDetails.forEach((drawing) => {
        // Add Drawing Row
        rows.push([
            `${srNo}`,
            drawing.dno,
            '', // ITEM is blank for the drawing row
            '', '', '', '', '', '', '', '', '', '', ''
        ]);

        drawing.partList.forEach((part) => {
            const weight = parseFloat(part.weight) || 0;
            const overhead = parseFloat(part.overhead) || 0;
            const totalWeight = weight + overhead;

            totalOverhead += overhead;
            totalWeightSum += weight;
            totalTWeightSum += totalWeight;

            rows.push([
                '', // Blank SR NO for parts
                '', // Blank DRG NO for parts
                part.partName,
                part.quantity,
                part.material,
                weight.toFixed(2),
                overhead.toFixed(2),
                totalWeight.toFixed(2),
                part.rate,
                part.labour,
                part.laserCut,
                part.primer,
                this.totalRate(part).toFixed(2),
                this.calculatePartTotal(part).toFixed(2),
            ]);
        });

        srNo++;
    });

    // Add Totals Row
    rows.push([
        '', '', 'TOTAL', '', '', totalWeightSum.toFixed(2), totalOverhead.toFixed(2),
        totalTWeightSum.toFixed(2), '', '', '', '', '', selectedRecord.totalAmount.toFixed(2)
    ]);

    // Adjust Column Widths
    const columnWidths = [
        10, // SR NO
        20, // DRG NO
        40, // ITEM
        15, // QTY
        20, // GRADE
        15, // WT
        15, // OH
        20, // T.WT
        20, // RATE
        20, // LABOUR
        15, // L/C
        15, // PRIMER
        20, // T.RATE
        25, // AMOUNT
    ];

    // Generate Table
    doc.autoTable({
        head: headers,
        body: rows,
        startY: 30,
        columnStyles: columnWidths.reduce((acc, width, index) => {
            acc[index] = { cellWidth: width };
            return acc;
        }, {}),
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
        didParseCell: function (data) {
            const { row, column } = data;
            const rowIndex = row.index; // Current row index

            // Check if it's the TOTAL row
            if (rowIndex === rows.length - 1) {
                // Make the entire TOTAL row bold
                data.cell.styles.fontStyle = 'bold';
            }

            // Alternatively, make only specific columns bold
            if (rowIndex === rows.length - 1 && (column.index === 2 || column.index === 13)) {
                // Column 2 (TOTAL label) and Column 13 (Total Amount)
                data.cell.styles.fontStyle = 'bold';
            }
        },
    });

    // Save PDF
    doc.save(`${rfqNumber}_quotation_report_revised.pdf`);
};



  public componentDidMount(): void {
    this.loadRFQNumbersFromSharePoint(); // Call the method here
  }
  
  private handleDownloadSecondPDF = async (rfqNumber: string) => {
    try {
      const web = new Web("https://skgroupenginering.sharepoint.com/sites/SalesManagement");
  
      // Fetch RFQ details from SharePoint
      const rfqItems = await web.lists
        .getByTitle("RFQList")
        .items.filter(`RFQNumber eq '${rfqNumber}'`)
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
      doc.text("Email: sonalienterprises89@rediffmail.com | Contact: 9960414239", 105, 22, {
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
        if (this.state.records[i].serialNumber === rfqNumber) {
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
        const { totalWeight,avgRate, totalAmount } = this.calculateDrawingTotals(drawing);
        const overhead = totalWeight * 0.1; // Overhead as 10% of total weight
  
        return [
          `${index + 1}`,
          drawing.dno,
          totalWeight.toFixed(2),
          overhead.toFixed(2),
          avgRate.toFixed(2),
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
      doc.save(`${rfqNumber}_Quotation_revised.pdf`);
    } catch (error) {
      console.error("Error generating PDF:", error);
      alert("Failed to generate PDF. Please try again.");
    }
  };

  private fetchHighestRevisionRecord = async (rfqNumber: string, web: Web): Promise<QuotationData | null> => {
    try {
      // Fetch records from QuotationList with Status 'Revised'
      const quotationItems = await web.lists
        .getByTitle("QuotationList")
        .items.filter(`RFQSerialNumber eq '${rfqNumber}' and Status eq 'Revised'`)
        .select("ID", "RFQSerialNumber", "RevisionNumber", "QuotationDate", "TotalAmount", "TotalRate", "TotalWeight", "Status", "Reason", "ApprovalDate")
        .get();
  
      // Fetch records from QuotationRevision with Status 'Revised'
      const revisionItems = await web.lists
        .getByTitle("QuotationRevision")
        .items.filter(`RFQSerialNumber eq '${rfqNumber}' and Statuss eq 'Revised'`)
        .select("ID", "RFQSerialNumber", "RevisionNumber", "RevisionDate", "TotalAmount", "TotalRate", "TotalWeight", "Statuss", "Reason", "ApprovalDate")
        .get();
  
      // Find the highest revision number across both lists
      let highestRevisionNumber = 0;
      let highestSource = "QuotationList"; // Track the source of the highest revision
  
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
  
      if (highestSource === "QuotationList") {
        // Fetch the record with the highest revision number from QuotationList
        const highestRevisionItem = quotationItems.find((item) =>
          parseInt(item.RevisionNumber || "0", 10) === highestRevisionNumber
        );
  
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
          };
        }
      } else if (highestSource === "QuotationRevision") {
        // Fetch the record with the highest revision number from QuotationRevision
        const highestRevisionItem = revisionItems.find((item) =>
          parseInt(item.RevisionNumber || "0", 10) === highestRevisionNumber
        );
  
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
          };
        }
      }
  
      return null;
    } catch (error) {
      console.error(`Error fetching highest revision record for RFQ ${rfqNumber}:`, error);
      return null;
    }
  };
  
  
  private handleSerialNumberChange = async (e: React.ChangeEvent<HTMLSelectElement>) => {
    const serialNumber = e.target.value;
  
    if (!serialNumber) return;
  
    try {
      const web = new Web("https://skgroupenginering.sharepoint.com/sites/SalesManagement"); 
  
      // Fetch the highest revision record for the selected serial number
      const highestRevisionRecord = await this.fetchHighestRevisionRecord(serialNumber, web);
  
      if (!highestRevisionRecord) {
        alert("No data found for the selected RFQ number with status 'Revised'.");
        return;
      }
  
      // Fetch other details related to the serial number
      const drawingDetails = await this.fetchDrawingAndPartDetailsBySerialNumber(serialNumber);
      const customerDetails = await this.fetchCustomerDetailsForSpecificRFQ(serialNumber);
  
      const currentDate = new Date().toISOString().split("T")[0];
  
      let existingRecordIndex = -1;
  
      // Find the index of the existing record
      for (let i = 0; i < this.state.records.length; i++) {
        if (this.state.records[i].serialNumber === serialNumber) {
          existingRecordIndex = i;
          break;
        }
      }
  
      const updatedRecords = [...this.state.records];
  
      if (existingRecordIndex >= 0) {
        // Update existing record
        updatedRecords[existingRecordIndex] = {
          ...updatedRecords[existingRecordIndex],
          drawingDetails,
          customerDetails: customerDetails || updatedRecords[existingRecordIndex].customerDetails,
          revisionNumber: highestRevisionRecord.revisionNumber.toString(),
          quotationDate: highestRevisionRecord.quotationDate,
          status: highestRevisionRecord.status,
          totalweight: highestRevisionRecord.totalWeight,
          totalAmount: highestRevisionRecord.totalAmount,
          revisionDate: currentDate,
        };
      } else {
        // Add new record
        const newRecord: IQuotationRecord = {
          id: updatedRecords.length + 1,
          serialNumber,
          rfqNumber: serialNumber,
          revisionNumber: highestRevisionRecord.revisionNumber.toString(),
          quotationDate: highestRevisionRecord.quotationDate,
          revisionDate: currentDate,
          status: highestRevisionRecord.status,
          drawingDetails,
          totalweight: highestRevisionRecord.totalWeight,
          totalRate: highestRevisionRecord.totalRate,
          totalAmount: highestRevisionRecord.totalAmount,
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
      alert("An error occurred while processing the serial number change. Please try again.");
    }
  };
  
  
  private handleDownloadBothPDFs = async (serialNumber: string) => {
    try {
      // Call the first PDF generation
      this.handleDownloadPDF(serialNumber);
  
      // Call the second PDF generation
      await this.handleDownloadSecondPDF(serialNumber);
  
      alert("Both PDFs downloaded successfully.");
    } catch (error) {
      console.error("Error downloading PDFs:", error);
      alert("Failed to download PDFs. Please try again.");
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
        const web = new Web("https://skgroupenginering.sharepoint.com/sites/SalesManagement")
  
        // Step 1: Update or Add DrawingList items
        for (const drawing of recordToConfirm.drawingDetails) {
          const existingDrawings = await web.lists
            .getByTitle("DrawingList")
            .items.filter(`RFQNumber eq '${recordToConfirm.rfqNumber}' and DrawingNumber eq '${drawing.dno}'`)
            .get();
            console.log(existingDrawings)
          if (existingDrawings.length > 0) {
            // Update existing drawing
            const itemId = existingDrawings[0].Id;
            await web.lists.getByTitle("DrawingList").items.getById(itemId).update({
              TotalWeight: drawing.totalWeight,
              AvgRate: drawing.avgRate,
              TotalAmount: drawing.totalAmount,
            });
          } 
        }
        console.log("first")
        // Step 2: Update QuotationList status
        const quotationItems = await web.lists
          .getByTitle("QuotationList")
          .items.filter(`RFQSerialNumber eq '${recordToConfirm.rfqNumber}'`)
          .select("Id")
          .get();
  
        if (quotationItems.length > 0) {
          const itemId = quotationItems[0].Id;
          await web.lists.getByTitle("QuotationList").items.getById(itemId).update({
            Status: "WorkingDone",
          });
        }
  
        // Step 3: Update RFQList status
        const rfqItems = await web.lists
          .getByTitle("RFQList")
          .items.filter(`RFQNumber eq '${recordToConfirm.rfqNumber}'`)
          .select("Id")
          .get();
  
        if (rfqItems.length > 0) {
          const itemId = rfqItems[0].Id;
          await web.lists.getByTitle("RFQList").items.getById(itemId).update({
            Status: "WorkingDone",
          });
        }
  
      // === Step 4: Update QuotationRevision list records to "WorkingDone" ===
      const revisionItems = await web.lists
        .getByTitle("QuotationRevision")
        .items.filter(`RFQSerialNumber eq '${recordToConfirm.rfqNumber}'`)
        .select("Id")
        .get();

      for (const item of revisionItems) {
        await web.lists.getByTitle("QuotationRevision").items.getById(item.Id).update({
          Statuss: "WorkingDone",
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
  
        alert("Drawing details and statuses updated successfully.");
      } catch (error) {
        console.error("Error updating drawing details and statuses:", error);
        alert("Failed to update drawing details and statuses. Please try again.");
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
      try {
      const rfqNumber = currentRecord.rfqNumber;


      // Step 1: Recalculate totals for the updated part details
      currentRecord.totalweight = currentRecord.drawingDetails.reduce((totalWeight, drawing) => {
        return totalWeight + drawing.partList.reduce((partWeight, part) => {
          return partWeight + this.totalWeight(part);
        }, 0);
      }, 0);

      currentRecord.totalRate = currentRecord.drawingDetails.reduce((totalRate, drawing) => {
        return totalRate + drawing.partList.reduce((partRate, part) => {
          return partRate + this.totalRate(part);
        }, 0);
      }, 0);

      currentRecord.totalAmount = currentRecord.drawingDetails.reduce((totalAmount, drawing) => {
        return totalAmount + drawing.partList.reduce((partAmount, part) => {
          return partAmount + this.calculatePartTotal(part);
        }, 0);
      }, 0);
  
        const web = new Web("https://skgroupenginering.sharepoint.com/sites/SalesManagement"); 
  

        // === Step 2: Update the PartList ===
          for (const drawing of currentRecord.drawingDetails) {
            for (const part of drawing.partList) {
              await this.updatePartDetailsInList(rfqNumber, part.partName, {
                material: part.material,
                weight: part.weight,
                overhead: part.overhead,
                rate: part.rate,
                labour: part.labour,
                laserCut: part.laserCut,
                primer: part.primer,
              });
            }
          }
          await web.lists.getByTitle("QuotationRevision").items.add({
            Title: `Revision for RFQ: ${rfqNumber}`,
            RFQSerialNumber: rfqNumber,
            RevisionNumber: parseInt(currentRecord.revisionNumber), // Convert to number   
            RevisionDate: new Date(currentRecord.revisionDate).toISOString(), // Convert to ISO format
            TotalWeight: currentRecord.totalweight, // Number
            TotalRate: currentRecord.totalRate, // Number
            TotalAmount: currentRecord.totalAmount,
            Statuss: "WorkingDone", // Number
          });
  
        // === Step 4: Update Local State ===
        const updatedRecords = records.map((record) =>
          record.id === currentRecord.id ? currentRecord : record
        );
  
        this.setState({
          records: updatedRecords,
          isEditing: false,
          currentRecord: null,
        });
  
        alert("Changes saved successfully! A new revision has been recorded.");
  
      } catch (error) {
        console.error("Error in saveEdit:", error.message || error);
        alert(`Error saving changes: ${error.message || "Unknown error occurred."}`);
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
        // Update part details within a drawing
        const updatedRecord = { ...prevState.currentRecord };
        const updatedPart = {
          ...updatedRecord!.drawingDetails[drawingIndex].partList[partIndex],
          [name]: value,
        };
  
        // Update the part details
        updatedRecord!.drawingDetails[drawingIndex].partList[partIndex] = updatedPart;
  
        // Recalculate drawing totals
        const { totalWeight, avgRate, totalAmount } = this.calculateDrawingTotals(updatedRecord!.drawingDetails[drawingIndex]);
        updatedRecord!.drawingDetails[drawingIndex].totalWeight = totalWeight;
        updatedRecord!.drawingDetails[drawingIndex].avgRate = avgRate;
        updatedRecord!.drawingDetails[drawingIndex].totalAmount = totalAmount;
  
        // Recalculate final totals
        updatedRecord!.totalweight = updatedRecord!.drawingDetails.reduce((sum, drawing) => sum + drawing.totalWeight, 0);
        updatedRecord!.totalAmount = updatedRecord!.drawingDetails.reduce((sum, drawing) => sum + drawing.totalAmount, 0);
  
        return { currentRecord: updatedRecord };
      } else if (name in prevState.currentRecord!) {
        // Update top-level fields (like revisionNumber)
        const updatedRecord = {
          ...prevState.currentRecord,
          [name]: value,
        };
  
        return { currentRecord: updatedRecord };
      }
  
      return null;
    });
  };
  
  
  
  private calculateDrawingTotals = (drawing: IDrawing): { totalWeight: number; avgRate: number; totalAmount: number } => {
    const totalWeight = drawing.partList.reduce((sum, part) => sum + this.totalWeight(part), 0) || 0;
    const totalRateSum = drawing.partList.reduce((sum, part) => sum + this.totalRate(part), 0) || 0;
    const totalAmount = drawing.partList.reduce((sum, part) => sum + this.calculatePartTotal(part), 0) || 0;
  
    const avgRate = drawing.partList.length > 0 ? totalRateSum / drawing.partList.length : 0;
    return { totalWeight, avgRate, totalAmount };
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
  
  private calculatePartTotal = (part: IPart): number => {
    return this.totalWeight(part) * this.totalRate(part);
  };
  



  public render(): React.ReactElement<IQuotationrevisionscreenProps> {
    const { records, isEditing, currentRecord, selectedSerialNumber, showSubmitConfirm, showDeleteConfirm } = this.state;

    return (
      <section className={styles.quotationrevisionscreen}>
      <h2>Quotation Revision Management Screen</h2>
      <div className={styles.dropdownBox}>
        <label htmlFor="serialNumber" className={styles.dropdownLabel}>
          Select Serial Number:
        </label>
        <select
          id="serialNumber"
          className={styles.selectDropdown}
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
              <p>Are you sure you have updated the data ?</p>
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
  <input
    type="date"
    name="quotationDate"
    value={currentRecord?.quotationDate || ""}
    onChange={this.handleChange}
  />
</label>

                <label>
                <label>
                  Revision Date:
                  <input type="date" name="revisionDate" value={currentRecord.revisionDate} onChange={this.handleChange} />
                </label>
                </label>
                <label>
                Revision Number:
                <input
                  type="text"
                  name="revisionNumber"
                  value={currentRecord?.revisionNumber || ""}
                  onChange={this.handleChange}
                />
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
                      <label>Material/Grade: <input type="text" name="material" value={part.material} onChange={(e) => this.handleChange(e, this.state.selectedDrawingIndex!, pIndex)} /></label>
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
            <strong>{part.partName}</strong> - material: {part.material}, 
            Weight: {part.weight}, Overhead: {part.overhead}, Rate: {part.rate}, 
            Labour: {part.labour}, Laser cut: {part.laserCut}, Primer: {part.primer}
            <br />
            <strong>Total Weight:</strong> {this.totalWeight(part).toFixed(2)}
            <br />
            <strong>Total Rate:</strong> {this.totalRate(part).toFixed(2)}
            <br />
            <strong>Part Total:</strong> {this.calculatePartTotal(part).toFixed(2)}
          </li>
        ))}
      </ul>
      <p>
        <strong>Drawing Total Weight:</strong> {(drawing.totalWeight || 0).toFixed(2)} <br />
        <strong>Average Rate:</strong> {(drawing.avgRate || 0).toFixed(2)} <br />
        <strong>Drawing Total Amount:</strong> {(drawing.totalAmount || 0).toFixed(2)}
      </p>
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
              Add Revised Data
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