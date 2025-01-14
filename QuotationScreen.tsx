import * as React from 'react';
import styles from './QuotationScreen.module.scss';
import jsPDF from 'jspdf';
import 'jspdf-autotable'
import { IQuotationScreenProps } from './IQuotationScreenProps';
import { Web }  from 'sp-pnp-js';

declare module 'jspdf' {
  interface jsPDF {
    autoTable: any;
    lastAutoTable?: { finalY: number };
  }
}

interface IPart {
  partName: string;
  material: string;
  grade: string;
  weight: string;
  overhead: string;
  rate: string;
  labour: string;
  laserCut: string;
  primer: string;
  quantity: string;
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
  quotationDate: Date | null;
  status: string;
  drawingDetails: IDrawing[];
  totalweight:number;
  totalAmount: number;
  avgRate:number;
  customerDetails?: ICustomerDetails; 
  rejectionReason?: string;
}

interface IQuotationScreenState {
  records: IQuotationRecord[];
  isEditing: boolean;
  selectedSerialNumber: string;
  currentRecord: IQuotationRecord | null;
  showSubmitConfirm: boolean;
  showDeleteConfirm: boolean;
  recordToConfirm: IQuotationRecord | null;
  selectedDrawingIndex: number | null;
  rfqNumbers: string[];
  materialList: IMaterial[];
}

interface ICustomerDetails {
  name: string;
  address: string;
  email: string;
  gstNumber: string;
  contactPerson: string;
  mobileNumbers: string[];
}

interface IMaterial {
  partNumber: string;
  material: string;
  rate: string;
  weight: string;
}

export default class QuotationScreen extends React.Component<IQuotationScreenProps, IQuotationScreenState> {
  constructor(props: IQuotationScreenProps) {
    super(props);
    this.state = {
      records: [],
      materialList: [],
      isEditing: false,
      selectedSerialNumber: "",
      currentRecord: null,
      showSubmitConfirm: false,
      showDeleteConfirm: false,
      recordToConfirm: null,
      selectedDrawingIndex: null,
      rfqNumbers: [],
    };
  }

  public componentDidMount(): void {
    this.loadRFQNumbersFromSharePoint();
    // this.fetchCustomerDetailsForRFQ();
  }
  private customerDetailsMap: Record<string, ICustomerDetails> = {};
  
  private loadRFQNumbersFromSharePoint = async (): Promise<void> => {
    try {
      const web = new Web("https://skgroupenginering.sharepoint.com/sites/SalesManagement") ;
    // Fetch only RFQs with Status as 'Todo'
    const items = await web.lists
      .getByTitle("RFQList")
      .items.filter("Status eq 'Todo'") 
      .select("RFQNumber")
      .get();

    const rfqNumbers = items.map((item: any) => item.RFQNumber);

    this.setState({ rfqNumbers });
  } catch (error) {
    console.error("Error loading RFQ numbers:", error);
  }
};
  
  private fetchCustomerDetailsForSpecificRFQ = async (rfqNumber: string): Promise<ICustomerDetails | null> => {
    try {
       const web = new Web("https://skgroupenginering.sharepoint.com/sites/SalesManagement"); // Add your site URL here
  
      // Fetch the RFQ details for the selected RFQNumber
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
  
      // Fetch the customer details for the associated CustomerName
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
  
  private fetchMaterialList = async (): Promise<void> => {
    try {
      const web = new Web("https://skgroupenginering.sharepoint.com/sites/SalesManagement");
  
      const items = await web.lists
        .getByTitle("MaterialList")
        .items.select("PartNumber", "Material", "Rate") // Include necessary fields
        .get();
  
      const materialList = items.map((item: any) => ({
        partNumber: item.PartNumber,
        material: item.Material,
        rate: item.Rate,
      }));
  
      console.log("Material List:", materialList); // Debug log
      this.setState({ materialList }); // Store as an array
    } catch (error) {
      console.error("Error fetching MaterialList:", error);
      alert("Failed to load material data.");
    }
  };
  
  
  
  private fetchDrawingAndPartDetailsBySerialNumber = async (rfqNumber: string): Promise<IDrawing[]> => {
    try {
      const web = new Web("https://skgroupenginering.sharepoint.com/sites/SalesManagement");
  
      // Fetch all drawings for the RFQ
      const drawingItems = await web.lists
        .getByTitle("DrawingList")
        .items.filter(`RFQNumber eq '${rfqNumber}'`)
        .select("DrawingNumber", "DrawingQuantity")
        .get();
  
      // Fetch all parts associated with the RFQ
      const partItems = await web.lists
        .getByTitle("PartList")
        .items.filter(`RFQNumber eq '${rfqNumber}'`)
        .select("DrawingNumber", "PartName", "Material", "Quantity", "Material", "Weight", "Overhead", "Rate", "Labour", "LaserCut", "Primer")
        .get();


    // Fetch MaterialList to match part details
    const materialItems = await web.lists
    .getByTitle("MaterialList")
    .items.select( "Weight", "Rate","Material")
    .get();

    const materialMap = new Map<string, any>();
    materialItems.forEach((material) => {
      materialMap.set(material.PartNumber, {
        weight: material.Weight,
        rate: material.Rate,
        material:material.Material,
      });
    });
  
      // Map drawing details and integrate part details
      const drawingDetails: IDrawing[] = drawingItems.map((drawing: any) => {
        // Filter parts for the current drawing
        const partsForDrawing = partItems.filter((part: any) => part.DrawingNumber === drawing.DrawingNumber);
  
        return {
          dlist: `Drawing ${drawing.DrawingNumber}`,
          dno: drawing.DrawingNumber,
          dquan: drawing.DrawingQuantity,
          partList: partsForDrawing.map((part: any) => {
            // Auto-populate fields if the part exists in the MaterialList
            const materialData = materialMap.get(part.PartName) || {};
  
            return {
              partName: part.PartName,
              material: part.Material,
              quantity: part.Quantity,
              weight: materialData.weight || part.Weight || "",
              overhead: part.Overhead || "",
              rate: materialData.rate || part.Rate || "",
              labour: part.Labour || "",
              laserCut: part.LaserCut || "",
              primer: part.Primer || "",
              width: materialData.width || "",
              thickness: materialData.thickness || "",
            };
          }),
        };
      });
  
      console.log("Drawing and Part Details Fetched", drawingDetails);
      return drawingDetails;
    } catch (error) {
      console.error("Error fetching drawing and part details:", error);
      alert("Failed to fetch drawing and part details. Please try again.");
      return [];
    }
  };


  private addQuotationToList = async (record: IQuotationRecord): Promise<void> => {
    try {
      const web = new Web("https://skgroupenginering.sharepoint.com/sites/SalesManagement"); 
      console.log("ADD Called")
      console.log("RFQ",record.serialNumber,record.revisionNumber,record.totalweight)
      
      const formattedDate = record.quotationDate
      ? record.quotationDate.toISOString().split('T')[0]
      : null;

      // Create a new item in the QuotationList
      await web.lists.getByTitle("QuotationList").items.add({
        RFQSerialNumber: record.serialNumber,
        QuotationDate: formattedDate, // Use formatted date string
        RevisionNumber: record.revisionNumber,
        TotalWeight: record.totalweight,
        TotalAmount: record.totalAmount,
      });
  
      alert("Quotation added successfully!");
    } catch (error) {
      console.error("Error adding quotation to the list:", error);
      alert("Failed to add quotation. Please try again.");
    }
  };

  
  // private addAdditionalPartDetails = async (
  //   rfqNumber: string,
  //   drawingNumber: string,
  //   partName: string,
  //   additionalDetails: Partial<IPart>
  // ): Promise<void> => {
  //   try {
  //     const web = new Web("https://skgroupenginering.sharepoint.com/sites/SalesManagement");
  
  //     // Check if the part exists
  //     const items = await web.lists
  //       .getByTitle("PartList")
  //       .items.filter(
  //         `RFQNumber eq '${rfqNumber}' and DrawingNumber eq '${drawingNumber}' and PartName eq '${partName}'`
  //       )
  //       .select("Id") // Only fetch the ID of the matching item
  //       .get();
  
  //     if (items.length === 0) {
  //       console.warn(
  //         `No matching part found for RFQNumber: ${rfqNumber}, DrawingNumber: ${drawingNumber}, and PartName: ${partName}`
  //       );
  //       alert(`No matching part found for RFQNumber: ${rfqNumber}, DrawingNumber: ${drawingNumber}, and PartName: ${partName}.`);
  //       return;
  //     }
  
  //     const itemId = items[0].Id; // Get the ID of the matching item
  
  //     // Add only the additional details to the existing item
  //     await web.lists.getByTitle("PartList").items.getById(itemId).update({
  //       Grade: additionalDetails.grade,
  //       Weight: additionalDetails.weight,
  //       Overhead: additionalDetails.overhead,
  //       Rate: additionalDetails.rate,
  //       Labour: additionalDetails.labour,
  //       LaserCut: additionalDetails.laserCut,
  //       Primer: additionalDetails.primer,
  //     });
  
  //     alert(`Additional details for "${partName}" added successfully!`);
  //   } catch (error) {
  //     console.error("Error adding additional details to the part:", error);
  //     alert("Failed to add additional details. Please try again.");
  //   }
  // };
  
  
  private handleSerialNumberChange = async (e: React.ChangeEvent<HTMLSelectElement>) => {
    const serialNumber = e.target.value;
  
    if (!serialNumber) return; // Exit if no RFQ is selected

    try {
      const web = new Web("https://skgroupenginering.sharepoint.com/sites/SalesManagement");
  
      await this.fetchMaterialList();

      // Fetch RFQ details
      const rfqItems = await web.lists
        .getByTitle("RFQList")
        .items.filter(`RFQNumber eq '${serialNumber}'`)
        .select("CustomerName", "Subject", "Date")
        .get();
  
      if (!rfqItems || rfqItems.length === 0) {
        alert("No RFQ data found for the selected RFQ number.");
        return;
      }
  
      const rfq = rfqItems[0];

       // Fetch "Reason" column from QuotationList
    const quotationItems = await web.lists
    .getByTitle("QuotationList")
    .items.filter(`RFQSerialNumber eq '${serialNumber}'`)
    .select("Reason","RevisionNumber") // Include Reason field
    .get();

    // Get the rejection reason (if any)
    const rejectionReason = quotationItems.length > 0 ? quotationItems[0].Reason : "";
    const revisionNumber = quotationItems.length > 0 ? quotationItems[0].RevisionNumber : "0";
      // Fetch drawing and part details
      const drawingDetails = await this.fetchDrawingAndPartDetailsBySerialNumber(serialNumber);
  
      // Populate customer details
      const customerDetails = await this.fetchCustomerDetailsForSpecificRFQ(serialNumber);
  
      // Auto-populate data into form fields
      const autoPopulatedRecord: IQuotationRecord = {
        id: this.state.records.length + 1,
        serialNumber,
        rfqNumber: serialNumber,
        revisionNumber,
        quotationDate: rfq.Date ? new Date(rfq.Date) : null,
        status: "Todo",
        rejectionReason,
        drawingDetails,
        totalweight: 0,
        totalAmount: 0,
        avgRate: 0,
        customerDetails: customerDetails || undefined,
      };
  
      // Calculate totals
      autoPopulatedRecord.totalweight = this.calculateTotalWeight(autoPopulatedRecord);
      autoPopulatedRecord.totalAmount = this.calculateTotalAmount(autoPopulatedRecord);
      autoPopulatedRecord.avgRate = this.calculateAverageRate(autoPopulatedRecord);
  
      // Manually find the index using `indexOf` equivalent
      const updatedRecords = [...this.state.records];
      let existingRecordIndex = -1;
      for (let i = 0; i < updatedRecords.length; i++) {
        if (updatedRecords[i].serialNumber === serialNumber) {
          existingRecordIndex = i;
          break;
        }
      }
  
      if (existingRecordIndex !== -1) {
        updatedRecords[existingRecordIndex] = autoPopulatedRecord;
      } else {
        updatedRecords.push(autoPopulatedRecord);
      }
  
      this.setState({
        selectedSerialNumber: serialNumber,
        currentRecord: autoPopulatedRecord,
        records: updatedRecords,
      });
    } catch (error) {
      console.error("Error auto-populating data:", error);
      alert("Failed to auto-populate data. Please try again.");
    }
  };
  
  // private handleEdit = (record: IQuotationRecord) => {
  //   this.setState({ isEditing: true, currentRecord: { ...record } ,selectedDrawingIndex: 0,});
  // };

  private handleEdit = (record: IQuotationRecord) => {
    const { materialList } = this.state;
  
    // Update drawingDetails with auto-filled Material and Rate
    const updatedDrawingDetails = record.drawingDetails.map((drawing) => {
      const updatedPartList = drawing.partList.map((part) => {
        // Use indexOf to find the index of the matching part
        const partIndex = materialList.map((m) => m.partNumber).indexOf(part.partName);
  
        // If the part exists, fetch its details; otherwise, provide defaults
        const materialDetails = partIndex !== -1 ? materialList[partIndex] : undefined;
  
        return {
          ...part,
          material: materialDetails?.material || part.material, // Populate Material
          rate: materialDetails?.rate || part.rate,             // Populate Rate
        };
      });
  
      return { ...drawing, partList: updatedPartList };
    });
  
    this.setState({
      isEditing: true,
      currentRecord: {
        ...record,
        drawingDetails: updatedDrawingDetails,
      },
      selectedDrawingIndex: 0, // Default to the first drawing
    });
  };
  
  
  
  private calculateAverageRate = (record: IQuotationRecord): number => {
    const totalRateSum = record.drawingDetails.reduce((sum, drawing) => {
      return sum + drawing.partList.reduce((partSum, part) => partSum + parseFloat(part.rate || '0'), 0);
    }, 0);
  
    const totalParts = record.drawingDetails.reduce((count, drawing) => count + drawing.partList.length, 0);
  
    return totalParts > 0 ? totalRateSum / totalParts : 0;
  };

  private handleDrawingChange = (e: React.ChangeEvent<HTMLSelectElement>) => {
    this.setState({ selectedDrawingIndex: parseInt(e.target.value) });
  };
  

  private handleSubmitConfirm = (record: IQuotationRecord) => {
    this.setState({ showSubmitConfirm: true, recordToConfirm: record });
  };

  private confirmSubmit = async () => {
    const { recordToConfirm, records } = this.state;
  
    if (recordToConfirm) {
      try {
        const web = new Web("https://skgroupenginering.sharepoint.com/sites/SalesManagement") // Replace with your SharePoint site URL
  
        // === Step 1: Update or Add DrawingList items ===
        for (const drawing of recordToConfirm.drawingDetails) {
          // Calculate totals for each drawing
          const { totalWeight, avgRate, totalAmount } = this.calculateDrawingTotals(drawing);
  
          // Check if the drawing already exists in DrawingList
          const existingDrawings = await web.lists
            .getByTitle("DrawingList")
            .items.filter(`RFQNumber eq '${recordToConfirm.rfqNumber}' and DrawingNumber eq '${drawing.dno}'`)
            .get();
  
          if (existingDrawings.length > 0) {
            // Update existing drawing
            const itemId = existingDrawings[0].Id;
            await web.lists.getByTitle("DrawingList").items.getById(itemId).update({
              TotalWeight: totalWeight,
              AvgRate: avgRate,
              TotalAmount: totalAmount,
            });
            console.log(`Updated DrawingList for DrawingNumber: ${drawing.dno}`);
          } else {
            // Add a new drawing record if it doesn't exist
            await web.lists.getByTitle("DrawingList").items.add({
              RFQNumber: recordToConfirm.rfqNumber,
              DrawingNumber: drawing.dno,
              TotalWeight: totalWeight,
              AvgRate: avgRate,
              TotalAmount: totalAmount,
            });
            console.log(`Added new DrawingList entry for DrawingNumber: ${drawing.dno}`);
          }
        }
  
        // === Step 2: Check if the record exists in QuotationList ===
        const existingQuotations = await web.lists
          .getByTitle("QuotationList")
          .items.filter(`RFQSerialNumber eq '${recordToConfirm.serialNumber}'`)
          .get();
  
        if (existingQuotations.length > 0) {
          // Update existing quotation
          const quotationId = existingQuotations[0].Id;
          await web.lists.getByTitle("QuotationList").items.getById(quotationId).update({
            QuotationDate: recordToConfirm.quotationDate?.toISOString().split("T")[0],
            RevisionNumber: recordToConfirm.revisionNumber,
            TotalWeight: recordToConfirm.totalweight,
            TotalAmount: recordToConfirm.totalAmount,
            Status: "WorkingDone",
          });
          console.log(`Updated QuotationList for SerialNumber: ${recordToConfirm.serialNumber}`);
        } else {
          // Add a new quotation if it doesn't exist
          await web.lists.getByTitle("QuotationList").items.add({
            RFQSerialNumber: recordToConfirm.serialNumber,
            QuotationDate: recordToConfirm.quotationDate?.toISOString().split("T")[0],
            RevisionNumber: recordToConfirm.revisionNumber,
            TotalWeight: recordToConfirm.totalweight,
            TotalAmount: recordToConfirm.totalAmount,
            Status: "WorkingDone",
          });
          console.log(`Added new QuotationList entry for SerialNumber: ${recordToConfirm.serialNumber}`);
        }
  
        // === Step 3: Update the RFQ status in RFQList ===
        const rfqItems = await web.lists
          .getByTitle("RFQList")
          .items.filter(`RFQNumber eq '${recordToConfirm.serialNumber}'`)
          .get();
  
        if (rfqItems.length > 0) {
          const rfqItemId = rfqItems[0].Id;
          await web.lists.getByTitle("RFQList").items.getById(rfqItemId).update({
            Status: "WorkingDone",
          });
          console.log(`Updated RFQList status to "WorkingDone" for RFQ: ${recordToConfirm.serialNumber}`);
        }
  
        // === Step 4: Update the local state ===
        const updatedRecords = records.map((rec) =>
          rec.id === recordToConfirm.id ? { ...rec, status: "WorkingDone" } : rec
        );
  
        this.setState({ records: updatedRecords, showSubmitConfirm: false, recordToConfirm: null });
        alert("Quotation added/updated and status updated successfully!");
  
      } catch (error) {
        console.error("Error during submit confirmation and status update:", error);
        alert("Failed to add or update the quotation. Please try again.");
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
      const updatedRecords = records.filter((record) => record.id !== recordToConfirm.id);
      this.setState({ records: updatedRecords, showDeleteConfirm: false, recordToConfirm: null });
    }
  };

  private cancelDelete = () => {
    this.setState({ showDeleteConfirm: false, recordToConfirm: null });
  };

  private saveEdit = async () => {
    const { currentRecord, records } = this.state;
  
    if (currentRecord) {
      const rfqNumber = currentRecord.rfqNumber;
  
      try {
        const web = new Web("https://skgroupenginering.sharepoint.com/sites/SalesManagement"); // Replace with your SharePoint site URL
  
        for (const drawing of currentRecord.drawingDetails) {
          for (const part of drawing.partList) {
            const partName = part.partName;
  
            // 1. Update MaterialList with new Weight, Rate, and Material
            await this.updateMaterialList(partName, part);
  
            // 2. Add new entry to RateDatePairList with Rate, Material, and Quotation Date
            await this.addToRateDatePairList(partName, part.material, part.rate, currentRecord.quotationDate);


          // 3. Update PartList with new Weight, Rate, and Material
          await this.updatePartList(rfqNumber, drawing.dno, partName, part);
          }
        }
  
        // Update the local state with calculated totals
        currentRecord.totalweight = this.calculateTotalWeight(currentRecord);
        currentRecord.totalAmount = this.calculateTotalAmount(currentRecord);
  
        const updatedRecords = records.map((record) =>
          record.id === currentRecord.id ? currentRecord : record
        );
  
        this.setState({
          records: updatedRecords,
          isEditing: false,
          currentRecord: null,
        });
  
        alert("Record saved successfully and updated in MaterialList and RateDatePairList!");
      } catch (error) {
        console.error("Error saving edits:", error);
        alert("Failed to save changes. Please try again.");
      }
    }
  };

  private updatePartList = async (
    rfqNumber: string,
    drawingNumber: string,
    partName: string,
    part: IPart
  ): Promise<void> => {
    try {
      const web = new Web("https://skgroupenginering.sharepoint.com/sites/SalesManagement");// Replace with your SharePoint site URL
  
      // Find the part in the PartList that matches RFQNumber, DrawingNumber, and PartName
      const items = await web.lists
        .getByTitle("PartList")
        .items.filter(
          `RFQNumber eq '${rfqNumber}' and DrawingNumber eq '${drawingNumber}' and PartName eq '${partName}'`
        )
        .select("Id")
        .get();
  
      if (items.length > 0) {
        const itemId = items[0].Id;
  
        // Update the existing part in PartList with new values
        await web.lists
          .getByTitle("PartList")
          .items.getById(itemId)
          .update({
            Weight: part.weight,
            Rate: part.rate,
            Material: part.material,
            Overhead: part.overhead,
            Labour: part.labour,
            LaserCut: part.laserCut,
            Primer: part.primer,
          });
  
        console.log(`PartList updated for PartName: ${partName}, DrawingNumber: ${drawingNumber}`);
      } else {
        console.warn(`No matching part found in PartList for PartName: ${partName}.`);
      }
    } catch (error) {
      console.error("Error updating PartList:", error);
      throw error;
    }
  };
  

  private updateMaterialList = async (partName: string, part: IPart): Promise<void> => {
    try {
      const web = new Web("https://skgroupenginering.sharepoint.com/sites/SalesManagement"); // Replace with your SharePoint site URL
  
      // Find the item in MaterialList that matches the PartName
      const items = await web.lists
        .getByTitle("MaterialList")
        .items.filter(`PartNumber eq '${partName}'`)
        .select("Id")
        .get();
  
      if (items.length > 0) {
        const itemId = items[0].Id;
  
        // Update the matching item with new values
        await web.lists
          .getByTitle("MaterialList")
          .items.getById(itemId)
          .update({
            Weight: part.weight,
            Rate: part.rate,
            Material: part.material,
          });
  
        console.log(`MaterialList updated for PartNumber: ${partName}`);
      } else {
        console.warn(`PartNumber ${partName} not found in MaterialList.`);
      }
    } catch (error) {
      console.error("Error updating MaterialList:", error);
      throw error;
    }
  };
  
  private addToRateDatePairList = async (
  partName: string,
  material: string,
  rate: string,
  quotationDate: Date | null
): Promise<void> => {
  try {
    if (!quotationDate) {
      console.warn("Quotation date is missing, skipping RateDatePairList update.");
      return;
    }

    const web = new Web("https://skgroupenginering.sharepoint.com/sites/SalesManagement");// Replace with your SharePoint site URL

    // Add a new record to RateDatePairList
    await web.lists
      .getByTitle("RateDatePairList")
      .items.add({
        PartNumber: partName,
        Material: material,
        Rate: rate,
        Date: quotationDate.toISOString(),
      });

    console.log(`RateDatePairList updated for PartNumber: ${partName}`);
  } catch (error) {
    console.error("Error adding to RateDatePairList:", error);
    throw error;
  }
};

  

  private cancelEdit = () => {
    this.setState({ isEditing: false, currentRecord: null });
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
  
  

  private handleChange = (
    e: React.ChangeEvent<HTMLInputElement | HTMLSelectElement>,
    drawingIndex?: number,
    partIndex?: number
  ) => {
    const { name, value } = e.target;
  
    // Ensure value matches the expected type
    let updatedValue: string;
  
    if (name === "quotationDate") {
      updatedValue = (e.target as HTMLInputElement).valueAsDate
        ? (e.target as HTMLInputElement).valueAsDate!.toISOString().split('T')[0] // Convert Date to string
        : ""; // Default to an empty string if null
    } else {
      updatedValue = value; // Default to string for other fields
    }
  
    // Update state synchronously
    this.setState((prevState) => {
      if (drawingIndex !== undefined && partIndex !== undefined) {
        const updatedRecord = { ...prevState.currentRecord };
        updatedRecord!.drawingDetails[drawingIndex].partList[partIndex] = {
          ...updatedRecord!.drawingDetails[drawingIndex].partList[partIndex],
          [name]: updatedValue,
        };
        return { currentRecord: updatedRecord };
      }
      return { currentRecord: { ...prevState.currentRecord!, [name]: updatedValue } };
    });
  };
  
  

  private calculateTotalWeight = (record: IQuotationRecord): number => {
    return record.drawingDetails.reduce((totalWeight, drawing) => {
      return totalWeight + drawing.partList.reduce((partWeight, part) => {
        return partWeight + this.totalWeight(part);
      }, 0);
    }, 0);
  };


  private calculateTotalAmount = (record: IQuotationRecord): number => {
    return record.drawingDetails.reduce((totalAmount, drawing) => {
      const { totalWeight, avgRate } = this.calculateDrawingTotals(drawing);
      return totalAmount + (avgRate * totalWeight);
    }, 0);
  };
  private calculateDrawingTotals = (drawing: IDrawing): { totalWeight: number; avgRate: number; totalAmount: number } => {
    let totalWeight = 0;
    let totalRateSum = 0;
    let partCount = drawing.partList.length;
  
    drawing.partList.forEach((part) => {
      const partWeight = this.totalWeight(part);
      const partRate = this.totalRate(part);
  
      totalWeight += partWeight;
      totalRateSum += partRate;
    });
  
    const avgRate = partCount > 0 ? totalRateSum / partCount : 0;
    const totalAmount = totalWeight * avgRate;
  
    return { totalWeight, avgRate, totalAmount };
  };
  
  
  
  private handleDownloadBothPDFs = async (serialNumber: string) => {
    console.log("Called")
    await Promise.all([
      this.handleDownloadPDF(serialNumber),
      this.handleDownloadSecondPDF(serialNumber),
    ]);
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
    doc.save(`${rfqNumber}_quotation_report.pdf`);
};


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
    doc.save(`${rfqNumber}_Quotation.pdf`);
  } catch (error) {
    console.error("Error generating PDF:", error);
    alert("Failed to generate PDF. Please try again.");
  }
};

  public render(): React.ReactElement<IQuotationScreenProps> {
    const { records, isEditing, currentRecord, selectedSerialNumber, showSubmitConfirm, showDeleteConfirm } = this.state;

    return (
      <section className={styles.quotationScreen}>
      <h2>Quotation Management Screen</h2>
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


        {/* Confirmation Dialogs */}
        {showSubmitConfirm && (
          <div className={styles.confirmOverlay}>
            <div className={styles.confirmBox}>
              <p>Are you sure you want to add this quotation and sent this quotation to approval ?</p>
              <button onClick={this.confirmSubmit} className={styles.confirmButton}>Yes</button>
              <button onClick={this.cancelSubmit} className={styles.cancelButton}>No</button>
            </div>
          </div>
        )}

        {showDeleteConfirm && (
          <div className={styles.confirmOverlay}>
            <div className={styles.confirmBox}>
              <p>Are you sure you want to delete this quotation ?</p>
              <button onClick={this.confirmDelete} className={styles.confirmButton}>Yes</button>
              <button onClick={this.cancelDelete} className={styles.cancelButton}>No</button>
            </div>
          </div>
        )}

        {/* Edit Dialog */}
        {isEditing && currentRecord && (
          <div className={styles.editOverlay}>
            <div className={styles.editBox}>
              <h3>Edit Record</h3>
              <div className={styles.scrollableContent}>
              <label>Quotation Date:
  <input
    type="date"
    name="quotationDate"
    value={
      currentRecord?.quotationDate
        ? currentRecord.quotationDate.toISOString().split('T')[0] // Format Date to YYYY-MM-DD
        : '' // Default to empty string if null
    }
    onChange={(e) => {
      const target = e.target as HTMLInputElement; // Ensure target is HTMLInputElement
      this.setState((prevState) => ({
        currentRecord: {
          ...prevState.currentRecord!,
          quotationDate: target.valueAsDate || null, // Use valueAsDate directly
        },
      }));
    }}
  />
</label>
                <label>Revision Number:
                  <input type="text" name="revisionNumber" value={currentRecord.revisionNumber} onChange={this.handleChange} />
                </label>
                {/* Drawing Selection Dropdown */}
      <label>Select Drawing:
        <select
          value={this.state.selectedDrawingIndex !== null ? this.state.selectedDrawingIndex : ""}
          onChange={this.handleDrawingChange}
          className={styles.drawingDropdown}
        >
          {currentRecord.drawingDetails.map((drawing, index) => (
            <option key={index} value={index}>
              {drawing.dlist} (Dno: {drawing.dno})
            </option>
          ))}
        </select>
      </label>
                  {this.state.selectedDrawingIndex !== null &&
                    currentRecord.drawingDetails[this.state.selectedDrawingIndex].partList.map((part, pIndex) => (
                      <div key={pIndex} className={styles.partEdit}>
                        <label>Part Name: <input type="text" value={part.partName} readOnly /></label>
                        <label>
  Grade/Material:
  <input
    type="text"
    name="material"
    value={part.material || ""}
    onChange={(e) => this.handleChange(e, this.state.selectedDrawingIndex!, pIndex)}
  />
</label>



                        <label>Weight: <input type="text" name="weight" value={part.weight} onChange={(e) => this.handleChange(e, this.state.selectedDrawingIndex!, pIndex)} /></label>
                        <label>Over Head: <input type="text" name="overhead" value={part.overhead} onChange={(e) => this.handleChange(e, this.state.selectedDrawingIndex!, pIndex)} /></label>
                        <label>Rate: <input type="text" name="rate" value={part.rate} onChange={(e) => this.handleChange(e, this.state.selectedDrawingIndex!, pIndex)} /></label>
                        <label>Labour: <input type="text" name="labour" value={part.labour} onChange={(e) => this.handleChange(e, this.state.selectedDrawingIndex!, pIndex)} /></label>
                        <label>Laser Cut: <input type="text" name="laserCut" value={part.laserCut} onChange={(e) => this.handleChange(e, this.state.selectedDrawingIndex!, pIndex)} /></label>
                        <label>Primer: <input type="text" name="primer" value={part.primer} onChange={(e) => this.handleChange(e, this.state.selectedDrawingIndex!, pIndex)} /></label>
                      </div>
                    ))}
                  {/* </div> */}
                {/* ))} */}
              </div>
              <button onClick={this.saveEdit} className={styles.saveButton}>Save</button>
              <button onClick={this.cancelEdit} className={styles.cancelButton}>Cancel</button>
            </div>
          </div>
        )}

        {/* Table */}
        <div className={styles.tableWrapper}>
          {records.length > 0 ? (
            <table className={styles.quotationTable}>
              <thead>
                <tr>
                  <th>RFQ Serial No</th>
                  <th>Customer Details</th>
                  <th style={{ minWidth: '250px' }}>Drawing & Part Details</th>
                  <th>Quotation date</th>
                  <th> Revison number </th>
                  <th>Total Weight</th>
                  <th>Total Amount</th>
                  <th>Status</th>
                  <th>Reject Reason</th>
                  <th>Actions</th>
                </tr>
              </thead>
              <tbody>
  {this.state.records.map((record) => (
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
  {record.drawingDetails.map((drawing, dIndex) => {
    const { totalWeight, avgRate, totalAmount } = this.calculateDrawingTotals(drawing);

    return (
      <div key={dIndex}>
        {/* Drawing Information */}
        <strong>{drawing.dlist || "Unnamed Drawing"}</strong> (Dno: {drawing.dno || "N/A"}, Quantity: {drawing.dquan || 0})
        <ul>
          {drawing.partList.map((part, pIndex) => {
            const partTotalWeight = this.totalWeight(part);
            const partTotalRate = this.totalRate(part);
            const partTotalAmount = partTotalWeight * partTotalRate;

            return (
              <li key={pIndex}>
                <strong>Part Name:</strong> {part.partName || "Unnamed Part"}<br />
                <strong>Material:</strong> {part.material || "N/A"}, Quantity: {part.quantity || "0"}<br />
                <strong>Grade:</strong> {part.material || "N/A"}<br />
                <strong>Weight:</strong> {part.weight || "0"}, <strong>Overhead:</strong> {part.overhead || "0"}<br />
                <strong>Total Weight:</strong> {partTotalWeight.toFixed(2)}<br />
                <strong>Rate:</strong> {part.rate || "0"}, Labour: {part.labour || "0"}, Laser Cut: {part.laserCut || "0"}, Primer: {part.primer || "0"}<br />
                <strong>Total Rate:</strong> {partTotalRate.toFixed(2)}<br />
                <strong>Total Amount:</strong> {partTotalAmount.toFixed(2)}
              </li>
            );
          })}
        </ul>
        {/* Display totals for the drawing */}
        <p>
          <strong>Total Weight:</strong> {totalWeight.toFixed(2)}<br />
          <strong>Average Rate:</strong> {avgRate.toFixed(2)}<br />
          <strong>Total Amount:</strong> {totalAmount.toFixed(2)}
        </p>
      </div>
    );
  })}
</td>


      <td>{record.quotationDate ? record.quotationDate.toISOString().split('T')[0] : ''}</td>
      <td>{record.revisionNumber}</td>
      <td>{record.totalweight}</td>
      <td>{record.totalAmount}</td>
      <td>{record.status}</td>
      <td>{record.rejectionReason || 'N/A'}</td> 
      <td>
        <button onClick={() => this.handleSubmitConfirm(record)} className={styles.submitButton}>Add Quotation</button>
        <button onClick={() => this.handleEdit(record)} className={styles.editButton}>Edit Quotation</button>
        <button onClick={() => this.handleDeleteConfirm(record)} className={styles.deleteButton}>Delete Quotation</button>
        <button onClick={() => this.handleDownloadBothPDFs(record.serialNumber)} className={styles.downloadButton}>Download PDF</button>    
      </td>
    </tr>
  ))}
</tbody>
            </table>
          ) : (
            <p>No records to display. Select a RFQ serial number to add a new Quotation.</p>
          )}
        </div>
      </section>
    );
  }
}