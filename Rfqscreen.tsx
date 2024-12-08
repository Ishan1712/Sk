import * as React from 'react';
import styles from './Rfqscreen.module.scss';
import { TextField, PrimaryButton, Dropdown, IDropdownOption, Dialog, DialogFooter, DialogType } from '@fluentui/react';
import { IRfqscreenProps } from './IRfqscreenProps';
import { Web }  from 'sp-pnp-js';

interface Rfq {
  Id: number; // Add this field
  rfqNumber: string;
  customerName: string;
  projectName: string;
  subject: string;
  date: Date;
  totalDrawings: number;
  contactPerson: string;
}

interface Drawing {
  drawingQuantity: string;
  drawingNumber: string;
  parts: Part[];
}

interface Part {
  material: string;
  quantity: string;
  partName: string;
}

interface IRfqScreenState {
  rfqNumber: string;
  customerName: string;
  projectName: string;
  subject: string; 
  date: Date | null;
  drawingQuantity: string;
  drawingNumber: string;
  partName: string;
  material: string;
  quantity: string;
  rfqList: Rfq[];
  drawingList: Drawing[];
  customerOptions: IDropdownOption[];
  currentPartList: Part[]; 
  isRfqFormVisible: boolean;
  isDrawingFormVisible: boolean;
  isMaterialFormVisible: boolean;
  isPartModalVisible: boolean;
  isEditModalVisible: boolean;
  isDeleteModalVisible: boolean;
  selectedRfq: Rfq | null;
  editCustomerName: string;
  editProjectName: string;
  editDate: Date | null;
  partNameOptions: IDropdownOption[]; 
}

export default class RfqScreen extends React.Component<IRfqscreenProps, IRfqScreenState> {
  private customerDetailsMap: Record<string, string> = {};
  constructor(props: IRfqscreenProps) {
    super(props);
    this.state = {
      rfqNumber: '',
      customerName: '',
      projectName: '',
      subject: '',
      date: null,
      drawingQuantity: '',
      drawingNumber: '',
      partName: '',
      material: '',
      quantity: '',
      rfqList: [],
      drawingList: [],
      currentPartList: [],
      customerOptions: [],
      isRfqFormVisible: false,
      isDrawingFormVisible: false,
      isMaterialFormVisible: false,
      isPartModalVisible: false, 
      isEditModalVisible: false, // Properly initialized
      isDeleteModalVisible: false, // Properly initialized
      selectedRfq: null, // Properly initialized
      editCustomerName: '',
      editProjectName: '',
      editDate: null,
      partNameOptions: [],
    };
  }

  private toggleRfqFormVisibility = (): void => {
    this.setState((prevState) => ({ isRfqFormVisible: !prevState.isRfqFormVisible }));
  };

  private toggleDrawingFormVisibility = (): void => {
    this.setState((prevState) => ({ isDrawingFormVisible: !prevState.isDrawingFormVisible }));
  };

  private togglePartModalVisibility = (): void => {
    this.setState((prevState) => ({ isPartModalVisible: !prevState.isPartModalVisible }));
  };

  private setRfqNumber = (value: string): void => this.setState({ rfqNumber: value });
  // private setCustomerName = (value: string): void => this.setState({ customerName: value });
  private setProjectName = (value: string): void => this.setState({ projectName: value });
  private setDate = (value: string): void => this.setState({ date: value ? new Date(value) : null });
  private setDrawingQuantity = (value: string): void => this.setState({ drawingQuantity: value });
  private setDrawingNumber = (value: string): void => this.setState({ drawingNumber: value });
  private setMaterial = (value: string): void => this.setState({ material: value });
  private setQuantity = (value: string): void => this.setState({ quantity: value });

  private addRfqToSharePoint = async (rfqNumber: string ,customerName: string,projectName: string,date: Date ,subject: string ): Promise<void> => {
    try {
      const web = new Web("https://skgroupenginering.sharepoint.com/sites/SalesManagement")
      await web.lists.getByTitle('RFQList').items.add({
        RFQNumber: rfqNumber,
        CustomerName: customerName,
        ProjectNumber: projectName,
        Subject: subject,
        Date: date,
      });
  
      alert("RFQ added successfully !");
    } catch (error) {
      console.error("Error adding RFQ :", error);
      alert("Failed to add RFQ. Please try again.");
    }
  };
  
  private loadCustomersFromSharePoint = async (): Promise<void> => {
    try {
      const web = new Web("https://skgroupenginering.sharepoint.com/sites/SalesManagement");
      const items = await web.lists
        .getByTitle('CustomerList')
        .items.select('CustomerName', 'ContactPerson') // Fetch ContactPerson column
        .getAll();
  
      const customerOptions = items.map((item: any) => ({
        key: item.CustomerName,
        text: item.CustomerName,
      }));
  
      // Map customer names to contact persons for reference
      const customerDetailsMap: Record<string, string> = {};
      items.forEach((item: any) => {
        customerDetailsMap[item.CustomerName] = item.ContactPerson || "N/A";
      });
  
      this.setState({ customerOptions });
      this.customerDetailsMap = customerDetailsMap; // Save to the state for later use
    } catch (error) {
      console.error('Error loading customers:', error);
    }
  };
  

  private addDrawingToSharePoint = async (
    rfqNumber: string,
    drawingNumber: string,
    drawingQuantity: string
  ): Promise<void> => {
    try {
      const web = new Web("https://skgroupenginering.sharepoint.com/sites/SalesManagement") 
      await web.lists.getByTitle('DrawingList').items.add({
        RFQNumber: rfqNumber,
        DrawingNumber: drawingNumber,
        DrawingQuantity: drawingQuantity,
      });

      alert('Drawing added successfully !');
    } catch (error) {
      console.error('Error adding Drawing :', error);
      alert('Failed to add Drawing. Please try again.');
    }
  };

  private addPartToSharePoint = async (
    rfqNumber: string,
    drawingNumber: string,
    partName : string ,
    material: string,
    quantity: string
  ): Promise<void> => {
    try {
      const web = new Web("https://skgroupenginering.sharepoint.com/sites/SalesManagement") 

      await web.lists.getByTitle('PartList').items.add({
        RFQNumber: rfqNumber,
        DrawingNumber: drawingNumber,
        PartName: partName,
        Material: material,
        Quantity: quantity,
      });

      alert('Part added successfully !');
    } catch (error) {
      console.error('Error adding Part :', error);
      alert('Failed to add Part. Please try again.');
    }
  };
  private getTotalDrawingsForRfq = async (rfqNumber: string): Promise<number> => {
    try {
     const web = new Web("https://skgroupenginering.sharepoint.com/sites/SalesManagement")
      const items = await web.lists
        .getByTitle('DrawingList')
        .items.filter(`RFQNumber eq '${rfqNumber}'`)
        .get();
      return items.length;
    } catch (error) {
      console.error('Error getting total drawings:', error);
      return 0;
    }
  };
  private loadRfqFromSharePoint = async (): Promise<void> => {
    try {
      const web = new Web("https://skgroupenginering.sharepoint.com/sites/SalesManagement");
      const items = await web.lists
        .getByTitle('RFQList')
        .items.select('Id', 'RFQNumber', 'CustomerName', 'ProjectNumber', 'Subject', 'Date')
        .getAll();
  
      const rfqList = await Promise.all(
        items.map(async (item: any) => {
          const totalDrawings = await this.getTotalDrawingsForRfq(item.RFQNumber);
          const contactPerson = this.customerDetailsMap[item.CustomerName] || "N/A"; // Use map for contact person
          return {
            Id: item.Id,
            rfqNumber: item.RFQNumber || '',
            customerName: item.CustomerName || '',
            projectName: item.ProjectNumber || '',
            subject: item.Subject || '',
            date: item.Date ? new Date(item.Date) : new Date(),
            totalDrawings,
            contactPerson, // Include Contact Person in RFQ data
          };
        })
      );
  
      this.setState({ rfqList });
    } catch (error) {
      console.error('Error loading RFQs:', error);
    }
  };

  // private loadDrawingFromSharePoint = async (): Promise<void> => {
  //   try {
  //     const web = new Web("https://skgroupenginering.sharepoint.com/sites/SalesManagement"); 
  
  //     const items = await web.lists
  //       .getByTitle('DrawingList')
  //       .items.select('RFQNumber', 'DrawingNumber', 'DrawingQuantity')
  //       .getAll();
  
  //     const drawingList: Drawing[] = items.map((item: any) => ({
  //       rfqNumber: item.RFQNumber || '', // Ensure RFQNumber is included
  //       drawingNumber: item.DrawingNumber || '',
  //       drawingQuantity: item.DrawingQuantity || '',
  //       parts: [], 
  //     }));
  
  //     this.setState({ drawingList });
  //   } catch (error) {
  //     console.error('Error loading Drawings:', error);
  //   }
  // };
  

  // private loadPartsFromSharePoint = async (): Promise<void> => {
  //   try {
  //     const web = new Web("https://skgroupenginering.sharepoint.com/sites/SalesManagement") 

  //     const items = await web.lists
  //       .getByTitle('PartList')
  //       .items.select('RFQNumber','DrawingNumber','PartName', 'Material', 'Quantity')
  //       .getAll();

  //     const partList: Part[] = items.map((item: any) => ({  
  //       partName: item.PartName || '',      
  //       material: item.Material || '',
  //       quantity: item.Quantity || '',
  //     }));

  //     this.setState({ currentPartList: partList });
  //   } catch (error) {
  //     console.error('Error loading Parts:', error);
  //     alert('Failed to load Parts. Please try again.');
  //   }
  // };

  private loadPartNamesFromSharePoint = async (): Promise<void> => {
    try {
      const web = new Web("https://skgroupenginering.sharepoint.com/sites/SalesManagement") 
      const items = await web.lists
        .getByTitle('MaterialList')
        .items.select('PartNumber') // Fetch PartNumber column
        .getAll();
  
      const partNameOptions = items.map((item: any) => ({
        key: item.PartNumber, // Use PartNumber as key
        text: item.PartNumber, // Display PartNumber in dropdown
      }));
  
      this.setState({ partNameOptions });
      console.log('Part names loaded from SharePoint:', partNameOptions);
    } catch (error) {
      console.error('Error loading part names from SharePoint:', error);
      alert('Failed to load part names. Please try again.');
    }
  };
  
  

  private addRfq = async (): Promise<void> => {
    const { rfqNumber,customerName, projectName,subject, date } = this.state;
  
    if (!rfqNumber || !customerName || !projectName || !date) {
      alert("Please fill all fields.");
      return;
    }
  
    try {
      // Call the function to add RFQ to SharePoint
      await this.addRfqToSharePoint( rfqNumber,customerName, projectName, date,subject);

      // const newPart: Rfq = {
      //   rfqNumber,
      //   customerName,
      //   projectName,
      //   date
      // };
  
      // Reset the form fields after successful addition
      this.setState({
        rfqNumber: '',
        customerName: '',
        projectName: '',
        subject: '', 
        date: null,
        isRfqFormVisible: false,
      });
    } catch (error) {
      console.error("Error adding RFQ:", error);
    }
  };
  
  private addDrawing = async (): Promise<void> => {
    const {
      rfqNumber,
      drawingQuantity,
      drawingNumber,
      drawingList,
      currentPartList,
    } = this.state;

    if (!drawingQuantity || !drawingNumber) {
      alert('Please fill all fields.');
      return;
    }
    try {
      await this.addDrawingToSharePoint(
        rfqNumber,
        drawingNumber,
        drawingQuantity
      );
    

      console.log(`Drawing ${drawingNumber} added for RFQ ${rfqNumber}`);
      const newDrawing: Drawing = {
        drawingQuantity,
        drawingNumber,
        parts: [...currentPartList],
      };

      this.setState({
        drawingList: [...drawingList, newDrawing],
        currentPartList: [],
        isDrawingFormVisible: false,
      });
    } catch (error) {
      console.error('Error adding Drawing:', error);
    }
  };


  private addPart = async (): Promise<void> => {
    const { rfqNumber,partName,material, quantity, drawingNumber, currentPartList } = this.state;

    if (!material || !partName|| !quantity || !drawingNumber) {
      alert('Please fill all fields.');
      return;
    }

    try {
      await this.addPartToSharePoint( rfqNumber,drawingNumber,partName ,material, quantity);

      const newPart: Part = {
        partName,
        material,
        quantity,
      };

      this.setState({
        currentPartList: [...currentPartList, newPart],
        partName:'',
        material: '',
        quantity: '',
        isPartModalVisible: false,
      });
    } catch (error) {
      console.error('Error adding Part:', error);
    }
  };

  // private submitRfq = (): void => {
  //   alert('RFQ Submitted');
  // };

  public componentDidMount(): void {
    this.loadRfqFromSharePoint();
    this.loadCustomersFromSharePoint();
    this.loadPartNamesFromSharePoint(); 
  }

  private showEditModal = (rfq: Rfq): void => {
    this.setState({
      isEditModalVisible: true,
      selectedRfq: rfq,
      editCustomerName: rfq.customerName,
      editProjectName: rfq.projectName,
      editDate: rfq.date,
    });
  };

  private hideEditModal = (): void => {
    this.setState({
      isEditModalVisible: false,
      selectedRfq: null,
    });
  };

  private saveEditedRfq = async (): Promise<void> => {
    const { selectedRfq, editCustomerName, editProjectName, editDate, rfqList } = this.state;
    if (!selectedRfq) return;
  
    try {
      const web = new Web("https://skgroupenginering.sharepoint.com/sites/SalesManagement") 
  
      // Update the RFQ in the SharePoint list
      await web.lists.getByTitle('RFQList').items.getById(selectedRfq.Id).update({
        CustomerName: editCustomerName,
        ProjectNumber: editProjectName,
        Subject: selectedRfq.subject,
        Date: editDate,
      });
  
      // Update the RFQ in the component state
      const updatedRfqList = rfqList.map((rfq) =>
        rfq.Id === selectedRfq.Id
          ? { ...rfq, customerName: editCustomerName, projectName: editProjectName,    subject: selectedRfq.subject, date: editDate! }
          : rfq
      );
  
      this.setState({
        rfqList: updatedRfqList,
        isEditModalVisible: false,
        selectedRfq: null,
      });
  
      alert('RFQ updated successfully!');
    } catch (error) {
      console.error('Error updating RFQ:', error);
      alert('Failed to update RFQ. Please try again.');
    }
  };
  

  private showDeleteModal = (rfq: Rfq): void => {
    this.setState({
      isDeleteModalVisible: true,
      selectedRfq: rfq,
    });
  };

  private hideDeleteModal = (): void => {
    this.setState({
      isDeleteModalVisible: false,
      selectedRfq: null,
    });
  };

  private deleteRfq = async (): Promise<void> => {
    const { selectedRfq, rfqList } = this.state;
    if (!selectedRfq) return;
  
    try {
      const web = new Web("https://skgroupenginering.sharepoint.com/sites/SalesManagement") 
  
      // Delete associated drawings
      const drawingsToDelete = await web.lists.getByTitle('DrawingList')
        .items.filter(`RFQNumber eq '${selectedRfq.rfqNumber}'`)
        .get();
  
      for (const drawing of drawingsToDelete) {
        await web.lists.getByTitle('DrawingList').items.getById(drawing.Id).delete();
      }
  
      // Delete associated parts
      const partsToDelete = await web.lists.getByTitle('PartList')
        .items.filter(`RFQNumber eq '${selectedRfq.rfqNumber}'`)
        .get();
  
      for (const part of partsToDelete) {
        await web.lists.getByTitle('PartList').items.getById(part.Id).delete();
      }
  
      // Delete the RFQ itself
      await web.lists.getByTitle('RFQList').items.getById(selectedRfq.Id).delete();
  
      // Update the state to remove the deleted RFQ from the UI
      const updatedRfqList = rfqList.filter((rfq) => rfq.Id !== selectedRfq.Id);
  
      this.setState({
        rfqList: updatedRfqList,
        isDeleteModalVisible: false,
        selectedRfq: null,
      });
  
      alert('RFQ and all associated data have been deleted successfully!');
    } catch (error) {
      console.error('Error deleting RFQ:', error);
      alert('Failed to delete RFQ. Please try again.');
    }
  };
  

  public render(): React.ReactElement {
    const {
      rfqNumber,
      customerName,
      projectName,
      date,
      drawingQuantity,
      drawingNumber,
      material,
      quantity,
      rfqList,
      // drawingList,
      // currentPartList,
      customerOptions,
      isRfqFormVisible,
      isDrawingFormVisible,
      isPartModalVisible,
      isEditModalVisible,
      isDeleteModalVisible,
      editCustomerName,
      editProjectName,
      editDate,
    } = this.state;

    const materialOptions: IDropdownOption[] = [
      { key: 'Aluminum', text: 'Aluminum' },
      { key: 'Steel', text: 'Steel' },
      { key: 'Plastic', text: 'Plastic' },
    ];

    return (
      <section className={styles.rfqScreen}>
        <h2>RFQ Management Screen</h2>
    
        <PrimaryButton
          className={styles.primaryButton}
          text={this.state.isRfqFormVisible ? "Close RFQ Form" : "Add RFQ"}
          onClick={this.toggleRfqFormVisibility}
        />
    
        {isRfqFormVisible && (
          <div className={styles.formContainer}>
            <TextField label="RFQ Number" value={rfqNumber} onChange={(e, newValue) => this.setRfqNumber(newValue || '')} />
            <Dropdown
              label="Customer Name"
              selectedKey={customerName}
              options={customerOptions}
              onChange={(e, option) =>
                this.setState({ customerName: option?.key as string })
              }
            />
            <TextField label="Project Number" value={projectName} onChange={(e, newValue) => this.setProjectName(newValue || '')} />
            <TextField label="Subject" value={this.state.subject} onChange={(e, newValue) => this.setState({ subject: newValue || '' })}/>

            <TextField label="Date" type="date" value={date ? date.toISOString().split('T')[0] : ''} onChange={(e, newValue) => this.setDate(newValue || '')} />
            <PrimaryButton className={styles.primaryButton} text={this.state.isDrawingFormVisible ? "Close Drawing Form" : "Add Drawing"} onClick={this.toggleDrawingFormVisibility} />
            <PrimaryButton className={styles.primaryButton} text="Submit RFQ" onClick={this.addRfq} />
          </div>
        )}
    
        {isDrawingFormVisible && (
          <div className={styles.drawingFormContainer}>
            <TextField label="Drawing Number" value={drawingNumber} onChange={(e, newValue) => this.setDrawingNumber(newValue || '')} />
            <TextField label="Drawing Quantity" value={drawingQuantity} onChange={(e, newValue) => this.setDrawingQuantity(newValue || '')} />
            <PrimaryButton className={styles.addPartButton} text="Add Part" onClick={this.togglePartModalVisibility} />
            <PrimaryButton className={styles.primaryButton} text="Submit Drawing" onClick={this.addDrawing} />
          </div>
        )}
    
        {/* Add Part Modal */}
        <Dialog
          hidden={!isPartModalVisible}
          onDismiss={this.togglePartModalVisibility}
          dialogContentProps={{
            type: DialogType.normal,
            title: 'Add Part',
          }}
        >
         <Dropdown
            label="Part Name"
            selectedKey={this.state.partName}
            onChange={(e, option) => this.setState({ partName: option?.key as string })}
            options={this.state.partNameOptions}
          />

          <Dropdown
            label="Material/grade"
            selectedKey={material} // Use the state variable to track the selected material
            onChange={(e, option) => this.setMaterial(option?.key as string)} // Update the state when selection changes
            options={materialOptions}
          />
          <TextField label="Quantity" value={quantity} onChange={(e, newValue) => this.setQuantity(newValue || '')} />
          <DialogFooter>
            <PrimaryButton text="Submit Part" onClick={this.addPart} />
            <PrimaryButton text="Cancel" onClick={this.togglePartModalVisibility} />
          </DialogFooter>
        </Dialog>
    
        <div className={styles.tableContainer}>
          <h3>RFQ List</h3>
          <table className={styles.rfqTable}>
            <thead>
              <tr>
                <th>RFQ Number</th>
                <th>Customer Name</th>
                <th>Contact Person</th>
                <th>Project Number</th>
                <th>Subject</th>
                <th>Date</th>
                <th>Total Drawings</th>
                <th>Actions</th>
              </tr>
            </thead>
            <tbody>
              {rfqList.map((rfq, index) => (
                <tr key={index}>
                  <td>{rfq.rfqNumber}</td>
                  <td>{rfq.customerName}</td>
                  <td>{rfq.contactPerson}</td>
                  <td>{rfq.projectName}</td>
                  <td>{rfq.subject}</td> {/* Display Subject */}
                  <td>{rfq.date.toISOString().split('T')[0]}</td>
                  <td>{rfq.totalDrawings}</td>
                  <td>
  <div style={{ display: "flex", gap: "10px", justifyContent: "center" }}>
    <PrimaryButton
      className={styles.editButton}
      text="Edit"
      onClick={() => this.showEditModal(rfq)}
    />
    <PrimaryButton
      className={styles.deleteButton}
      text="Delete"
      onClick={() => this.showDeleteModal(rfq)}
    />
  </div>
</td>

                </tr>
              ))}
            </tbody>
          </table>
          {/* Edit Modal */}
        <Dialog
          hidden={!isEditModalVisible}
          onDismiss={this.hideEditModal}
          dialogContentProps={{
            type: DialogType.largeHeader,
            title: 'Edit RFQ',
          }}
        >
          <TextField
            label="Customer Name"
            value={editCustomerName}
            onChange={(e, newValue) => this.setState({ editCustomerName: newValue || '' })}
          />
          <TextField
            label="Project Number"
            value={editProjectName}
            onChange={(e, newValue) => this.setState({ editProjectName: newValue || '' })}
          />
          <TextField
            label="Subject"
            value={this.state.selectedRfq?.subject || ''} // Bind to the selected RFQ's subject
            onChange={(e, newValue) =>
            this.setState((prevState) => ({
            selectedRfq: { ...prevState.selectedRfq!, subject: newValue || '' },
            }))
          }
         />
          <TextField
            label="Date"
            type="date"
            value={editDate ? editDate.toISOString().split('T')[0] : ''}
            onChange={(e, newValue) => this.setState({ editDate: newValue ? new Date(newValue) : null })}
          />
          <DialogFooter>
            <PrimaryButton text="Save" onClick={this.saveEditedRfq} />
            <PrimaryButton text="Cancel" onClick={this.hideEditModal} />
          </DialogFooter>
        </Dialog>
        {/* Delete Modal */}
        <Dialog
          hidden={!isDeleteModalVisible}
          onDismiss={this.hideDeleteModal}
          dialogContentProps={{
            type: DialogType.largeHeader,
            title: 'Confirm Delete',
            subText: 'Are you sure you want to delete this RFQ?',
          }}
        >
          <DialogFooter>
            <PrimaryButton text="Yes" onClick={this.deleteRfq} />
            <PrimaryButton text="No" onClick={this.hideDeleteModal} />
          </DialogFooter>
        </Dialog>
        </div>
      </section>
    );
  }
}