import * as React from 'react';
import { IDropdownOption } from '@fluentui/react/lib/Dropdown';
import { PrimaryButton } from '@fluentui/react/lib/Button';
import { TextField } from '@fluentui/react/lib/TextField';
import { Dropdown } from '@fluentui/react/lib/Dropdown';
import styles from './Masterscreen.module.scss';
import type { IMasterscreenProps } from './IMasterscreenProps';
import { Web }  from 'sp-pnp-js';

interface Material {
  id?: number; // Include ID for edit functionality
  partNumber: string;
  length :string;
  width: string;
  thickness : string;
  material: string;
  weight: string;
  rate: string;
  date: Date;
}

interface IMasterscreenState {
  partNumber: string;
  length: string;
  width: string;
  thickness: string;
  material: string;
  weight: string;
  rate: string;
  date: Date | null;
  materials: Material[];
  isFormVisible: boolean;
  editingIndex: number | null; // Track index for editing
  showDeleteModal: boolean; // For delete confirmation modal
  currentMaterial: Material | null; // To store the material being deleted
}

export default class Masterscreen extends React.Component<IMasterscreenProps, IMasterscreenState> {
  constructor(props: IMasterscreenProps) {
    super(props);
    this.state = {
      partNumber: '',
      length: '',
      width: '',
      thickness:'' ,
      material: '',
      weight: '',
      rate: '',
      date: null,
      materials: [],
      isFormVisible: false,
      editingIndex: null, // Initialize with null
      showDeleteModal: false, // Initialize delete modal as hidden
      currentMaterial: null, // Initialize with null
    };
  }

  private toggleFormVisibility = (): void => {
    this.setState((prevState) => ({ isFormVisible: !prevState.isFormVisible }));
  };

  private setPartNumber = (value: string): void => this.setState({ partNumber: value });
  private setLength = (value: string): void => this.setState({ length: value });
  private setWidth = (value: string): void => this.setState({ width: value });
  private setThickness = (value: string): void => this.setState({ thickness: value });
  private setMaterial = (value: string): void => this.setState({ material: value });
  private setWeight = (value: string): void => this.setState({ weight: value });
  private setRate = (value: string): void => this.setState({ rate: value });
  private setDate = (value: string): void => this.setState({ date: value ? new Date(value) : null });

  private addPartNumberToSharePoint = async (partNumber: string, length: string, width: string, thickness: string, material: string, weight: string, rate : string, date : Date): Promise<void> => {
    try {
      const web = new Web("https://skgroupenginering.sharepoint.com/sites/SalesManagement") // Replace with your SharePoint site URL
      await web.lists.getByTitle('MaterialList').items.add({
        PartNumber: partNumber, 
        Length: length,
        Width: width,
        Thickness: thickness,
        Material: material, 
        Weight : weight ,
        Rate : rate,
        Date : date
        // Assuming "Title" column stores part numbers
      });
      alert("Data added successfully!");
      console.log('Part number and length and width and thickness and material and weight and rate and date added successfully:', partNumber, length , width, thickness, material, weight, rate,date);
    } catch (error) {
      alert("ERROR occurred. Data is not added.");
      console.error('Error adding part number to SharePoint list:', error);
      alert('Failed to add part number. Please try again.');
    }
  };

  private loadMaterialsFromSharePoint = async (): Promise<void> => {
    try {
      const web = new Web("https://skgroupenginering.sharepoint.com/sites/SalesManagement")  // Replace with your SharePoint site URL
  
      // Fetch all items from the MaterialList
      const items = await web.lists.getByTitle('MaterialList').items.select('ID','PartNumber', 'Length', 'Width', 'Thickness', 'Material', 'Weight', 'Rate', 'Date').getAll();
  
      // Map SharePoint data to the Material type
      const materials = items.map((item: any) => ({
        id: item.ID, // Map the ID field
        partNumber: item.PartNumber || '', // Ensure fallback if the column is missing or empty
        length: item.Length || '',
        width: item.Width || '',
        thickness: item.Thickness || '',
        material: item.Material || '',
        weight: item.Weight || '',
        rate: item.Rate || '',
        date: item.Date ? new Date(item.Date) : new Date(), // Convert date to JavaScript Date object
      }));
  
      // Update the materials state
      this.setState({ materials });
      // alert("Material loaded successfully!");
      console.log("Materials loaded from SharePoint:", materials);
    } catch (error) {
      console.error("Error loading materials from SharePoint:", error);
      alert("Failed to load materials. Please try again.");
    }
  };
  public componentDidMount(): void {
    this.loadMaterialsFromSharePoint();
  }  

  private editMaterialInSharePoint = async (material: Material): Promise<void> => {
    if (!material.id) {
      alert("Invalid material ID.");
      return;
    }
  
    try {
      const web = new Web("https://skgroupenginering.sharepoint.com/sites/SalesManagement")  // Replace with your SharePoint site URL
      await web.lists.getByTitle('MaterialList').items.getById(material.id).update({
        PartNumber: material.partNumber,
        Length: material.length,
        Width: material.width,
        Thickness: material.thickness,
        Material: material.material,
        Weight: material.weight,
        Rate: material.rate,
        Date: material.date.toISOString(), // Ensure proper date format
      });
  
      alert("Material updated successfully!");
      console.log(`Material with ID ${material.id} updated.`);
    } catch (error) {
      console.error("Error updating material in SharePoint:", error);
      alert("Failed to update material. Please try again.");
    }
  };
  
  private deleteMaterialInSharePoint = async (materialId: number): Promise<void> => {
    if (!materialId) {
      alert("Invalid material ID.");
      return;
    }
  
    try {
      const web = new Web("https://skgroupenginering.sharepoint.com/sites/SalesManagement")  // Replace with your SharePoint site URL
      await web.lists.getByTitle('MaterialList').items.getById(materialId).delete();
      alert("Material deleted successfully!");
      console.log(`Material with ID ${materialId} deleted from SharePoint.`);
    } catch (error) {
      console.error("Error deleting material in SharePoint:", error);
      alert("Failed to delete material. Please try again.");
    }
  };
  
private addMaterial = async (): Promise<void> => {
  const { partNumber, length, width, thickness, material, weight, rate, date, materials, editingIndex } = this.state;

  if (!date) return;

  const newMaterial: Material = { partNumber, length, width, thickness, material, weight, rate, date };

  if (editingIndex !== null) {
    const updatedMaterials = [...materials];
    const materialToEdit = { ...newMaterial, id: materials[editingIndex].id }; // Add ID for SharePoint update

    try {
      await this.editMaterialInSharePoint(materialToEdit); // Update material in SharePoint
      updatedMaterials[editingIndex] = materialToEdit; // Update local state
      this.setState({ materials: updatedMaterials, editingIndex: null });
    } catch (error) {
      console.error("Error updating material in SharePoint:", error);
    }
  } else {
    try {
      await this.addPartNumberToSharePoint(partNumber, length, width, thickness, material, weight, rate, date);
      await this.loadMaterialsFromSharePoint();
    } catch (error) {
      console.error('Error adding material to SharePoint:', error);
    }
  }

  // Reset form fields
  this.setState({
    partNumber: '',
    length: '',
    width: '',
    thickness: '',
    material: '',
    weight: '',
    rate: '',
    date: null,
    isFormVisible: false,
  });
};

  private editMaterial = (index: number): void => {
    const material = this.state.materials[index];

    this.setState({
      partNumber: material.partNumber,
      length:material.length,
      width: material.width,
      thickness: material.thickness,
      material: material.material,
      weight: material.weight,
      rate: material.rate,
      date: material.date,
      isFormVisible: true,
      editingIndex: index, // Set editing index
    });
  };

  private deleteMaterial = (material: Material): void => {
    this.setState({ showDeleteModal: true, currentMaterial: material });
  };
  
  private confirmDelete = async (): Promise<void> => {
    const { currentMaterial, materials } = this.state;
  
    if (currentMaterial && currentMaterial.id) {
      try {
        await this.deleteMaterialInSharePoint(currentMaterial.id); // Delete from SharePoint
        const updatedMaterials = materials.filter((m) => m.id !== currentMaterial.id); // Remove locally
        this.setState({
          materials: updatedMaterials,
          showDeleteModal: false,
          currentMaterial: null,
        });
      } catch (error) {
        console.error("Error confirming delete:", error);
      }
    }
  };
  
  private cancelDelete = (): void => {
    this.setState({ showDeleteModal: false, currentMaterial: null });
  };
  

  public render(): React.ReactElement<IMasterscreenProps> {
    const { partNumber, length, width,thickness, material, weight, rate, date, materials, isFormVisible,showDeleteModal, } = this.state;

    const materialOptions: IDropdownOption[] = [
      { key: 'MM', text: 'MM' },
      { key: 'SS', text: 'SS' },
      { key: 'AL', text: 'AL' }
    ];

    return (
      <section className={styles.masterscreen}>
        <h2>Material Master Screen</h2>

        <PrimaryButton className={styles.PrimaryButton} text ={isFormVisible ? "Cancel" : "Add Material"} onClick={this.toggleFormVisibility} />

        {isFormVisible && (
          <div className={styles.formContainer}>
            <TextField label="Part Number" value={partNumber} onChange={(e, newValue) => this.setPartNumber(newValue || '')} />
            <TextField label="Length" value={length} onChange={(e, newValue) => this.setLength(newValue || '')}/>
            <TextField label="Width" value={width} onChange={(e, newValue) => this.setWidth(newValue || '')}/>
            <TextField label="Thickness" value={thickness} onChange={(e, newValue) => this.setThickness(newValue || '')}/>
            <Dropdown label="Material" selectedKey={material} onChange={(e, option) => this.setMaterial(option?.key as string)} options={materialOptions} />
            <TextField label="Weight" value={weight} onChange={(e, newValue) => this.setWeight(newValue || '')}/>
            <TextField label="Rate" value={rate} onChange={(e, newValue) => this.setRate(newValue || '')} />
            <TextField label="Date" type="date" value={date ? date.toISOString().split('T')[0] : ''} onChange={(e, newValue) => this.setDate(newValue || '')} />
            <PrimaryButton text={this.state.editingIndex !== null ? "Update Material" : "Save Material Details"} onClick={this.addMaterial} />
          </div>
        )}

        <h3>Materials List</h3>
        <div className={styles.materialTableContainer}>
        <table className={styles.materialTable}>
          <thead>
            <tr>
              <th>Part Number</th>
              <th>Length</th>
              <th>Width</th>
              <th>Thickness</th>
              <th>Material</th>
              <th>Weight</th>
              <th>Rate</th>
              <th>Date</th>
              <th>Actions</th>
            </tr>
          </thead>          
          <tbody>
            {materials.length > 0 ? (
              materials.map((item, index) => (
                <tr key={index}>
                  <td>{item.partNumber}</td>
                  <td>{item.length}</td>
                  <td>{item.width}</td>
                  <td>{item.thickness}</td>
                  <td>{item.material}</td>
                  <td>{item.weight}</td>
                  <td>{item.rate}</td>
                  <td>{item.date.toISOString().split('T')[0]}</td>
                  <td>
                    <div className={styles.actions}>
                      <button  className={styles.editButton} onClick={() => this.editMaterial(index)}>Edit</button>
                      <button className={styles.deleteButton} onClick={() => this.deleteMaterial(materials[index])}>Delete</button>
                    </div>
                  </td>
                </tr>
              ))
            ) : (
              <tr>
                <td colSpan={7} style={{ textAlign: 'center' }}>No materials added</td>
              </tr>
            )}
          </tbody>
        </table>
      </div>
        {/* Delete Confirmation Modal */}
        {showDeleteModal && (
          <div className={styles.modal}>
            <div className={styles.modalContent}>
              <p className={styles.modalMessage}>Are you sure you want to delete this material?</p>
              <div className={styles.buttonGroup}>
                <button onClick={this.confirmDelete} className={styles.yesButton}>Yes</button>
                <button onClick={this.cancelDelete} className={styles.noButton}>No</button>
              </div>
            </div>
          </div>
        )}
      </section>
    );
  }
}