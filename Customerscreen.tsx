import * as React from 'react';
import styles from './Customerscreen.module.scss';
import { ICustomerscreenProps } from './ICustomerscreenProps';
import { TextField, PrimaryButton, Modal } from '@fluentui/react';
import { Web } from 'sp-pnp-js';

interface Customer {
  id?: number; // Include ID for edit functionality
  name: string;
  address: string;
  mobileNumbers: string[];
  email: string;
  gstNumber: string;
  contactPerson: string;
}

interface ICustomerscreenState {
  customers: Customer[];
  newCustomer: Customer;
  editingIndex: number | null;
  searchQuery: string;
  isFormVisible: boolean;
  isDeleteModalVisible: boolean;
  deleteIndex: number | null;
}

export default class Customerscreen extends React.Component<ICustomerscreenProps, ICustomerscreenState> {
  constructor(props: ICustomerscreenProps) {
    super(props);
    this.state = {
      customers: [],
      newCustomer: {
        id: undefined, // For tracking SharePoint item ID
        name: '',
        address: '',
        mobileNumbers: [''],
        email: '',
        gstNumber: '',
        contactPerson: '',
      },
      editingIndex: null,
      searchQuery: '',
      isFormVisible: false,
      isDeleteModalVisible: false,
      deleteIndex: null,
    };
  }

  private toggleFormVisibility = (): void => {
    this.setState((prevState) => ({
      isFormVisible: !prevState.isFormVisible,
      editingIndex: null,
      newCustomer: {
        id: undefined,
        name: '',
        address: '',
        mobileNumbers: [''],
        email: '',
        gstNumber: '',
        contactPerson: '',
      },
    }));
  };

  private addCustomerToSharePoint = async (name: string, address : string, email:string, gstNumber:string,contactPerson:string, mobileNumbers:string[] ): Promise<void> => {
    try {
      const web = new Web("https://skgroupenginering.sharepoint.com/sites/SalesManagement") // Replace with your SharePoint site URL
      const mobileNumbersString = mobileNumbers.join(", ");
      await web.lists.getByTitle('CustomerList').items.add({
        CustomerName : name,
        Address : address,
        Email: email,
        GSTNumber :gstNumber,
        ContactPerson:contactPerson,
        MobileNumber: mobileNumbersString,
      });
      alert("Customer added successfully!");
      console .log(' ustomer name and address added successfully',name, address,email,gstNumber,contactPerson,mobileNumbersString)
    } catch (error) {
      console.error("Error adding customer to SharePoint:", error);
      if (error.data) {
        console.error("Error details:", error.data);
      }
      alert("ERROR occurred. Data is not added.");
    }
  };
  private loadCustomersFromSharePoint = async (): Promise<void> => {
    try {
      const web = new Web("https://skgroupenginering.sharepoint.com/sites/SalesManagement")  // Replace with your SharePoint site URL
      const items = await web.lists.getByTitle('CustomerList').items.select('ID','CustomerName', 'Address','Email','GSTNumber','ContactPerson','MobileNumber').getAll();
  
      // Map SharePoint data to the Customer type
      const customers = items.map((item: any) => ({
        id: item.ID, // Map the Id field
        name: item.CustomerName || '', // Assuming the "Customer Name" is stored in the "Title" field
        address:item.Address ||  '',
        email: item.Email || '',
        gstNumber:item.GSTNumber || '',
        contactPerson: item.ContactPerson || '',
        mobileNumbers: item.MobileNumber ? item.MobileNumber.split(", ") : [], // Split the string into an array
      }));
  
      // Update the state
      this.setState({ customers });
      // alert("Custmer details loaded successfully");
      console.log("Customers loaded from SharePoint:", customers);
    } catch (error) {
      alert("Failed to load customers. Please try again.");
      console.error("Error loading customers from SharePoint:", error);
    }
  };

  public componentDidMount(): void {
    this.loadCustomersFromSharePoint();
  }
  
  private addOrUpdateCustomer1 = async (): Promise<void> => {
    const { customers, newCustomer, editingIndex } = this.state;

    if (!newCustomer.name) {
        alert("Customer Name is required.");
        return;
    }

    if (editingIndex !== null) {
        // Updating an existing customer
        try {
            await this.editCustomerInSharePoint(newCustomer); // Call editCustomerInSharePoint
            customers[editingIndex] = newCustomer; // Update the local state
            this.setState({ customers, editingIndex: null });
            alert("Customer updated successfully!");
        } catch (error) {
            console.error("Error updating customer:", error);
        }
    } else {
        // Adding a new customer
        try {
            await this.addCustomerToSharePoint(
                newCustomer.name,
                newCustomer.address,
                newCustomer.email,
                newCustomer.gstNumber,
                newCustomer.contactPerson,
                newCustomer.mobileNumbers
            );
            this.setState({ customers: [...customers, newCustomer] });
        } catch (error) {
            console.error("Error adding customer:", error);
        }
    }

    this.resetForm();
};

private deleteCustomerFromSharePoint = async (id: number): Promise<void> => {
  try {
    const web = new Web("https://skgroupenginering.sharepoint.com/sites/SalesManagement")  // Replace with your SharePoint site URL
    await web.lists.getByTitle('CustomerList').items.getById(id).delete();
    alert("Customer deleted successfully!");
    console.log(`Customer with ID ${id} deleted successfully.`);
  } catch (error) {
    alert("Failed to delete customer. Please try again.");
    console.error("Error deleting customer from SharePoint:", error);
  }
};

  
  private editCustomerInSharePoint = async (customer: Customer): Promise<void> => {
    if (!customer.id) {
      alert("Invalid customer ID.");
      return;
    }

    try {
      const web = new Web("https://skgroupenginering.sharepoint.com/sites/SalesManagement")  // Replace with your SharePoint site URL
      const mobileNumbersString = customer.mobileNumbers.join(", ");
      await web.lists.getByTitle('CustomerList').items.getById(customer.id).update({
        CustomerName: customer.name,
        Address: customer.address,
        Email: customer.email,
        GSTNumber: customer.gstNumber,
        ContactPerson: customer.contactPerson,
        MobileNumber: mobileNumbersString,
      });
      alert("Customer updated successfully!");
      console.log("Customers updated from SharePoint:");
    } catch (error) {
      alert("Failed to update customer.");
      console.error("Error updating customer in SharePoint:", error);
    }
  };

  private handleInputChange = (field: keyof Customer, value: string) => {
    this.setState((prevState) => ({
      newCustomer: { ...prevState.newCustomer, [field]: value },
    }));
  };

  private resetForm = () => {
    this.setState({
      newCustomer: {
        name: '',
        address: '',
        mobileNumbers: [''],
        email: '',
        gstNumber: '',
        contactPerson: '',
      },
      isFormVisible: false,
    });
  };

  private editCustomer = (index: number) => {
    this.setState({
      newCustomer: { ...this.state.customers[index] },
      editingIndex: index,
      isFormVisible: true,
    });
  };

  private showDeleteModal = (index: number): void => {
    this.setState({
      isDeleteModalVisible: true,
      deleteIndex: index,
    });
  };

  private closeDeleteModal = (): void => {
    this.setState({
      isDeleteModalVisible: false,
      deleteIndex: null,
    });
  };

  private deleteCustomer = async (): Promise<void> => {
    const { customers, deleteIndex } = this.state;
    if (deleteIndex !== null) {
      const customerToDelete = customers[deleteIndex];
  
      if (customerToDelete.id) {
        try {
          // Call the SharePoint delete function
          await this.deleteCustomerFromSharePoint(customerToDelete.id);
  
          // Remove the customer from the state
          const updatedCustomers = customers.filter((_, i) => i !== deleteIndex);
          this.setState({
            customers: updatedCustomers,
            isDeleteModalVisible: false,
            deleteIndex: null,
          });
        } catch (error) {
          console.error("Error deleting customer:", error);
        }
      } else {
        alert("Invalid customer ID. Cannot delete.");
      }
    }
  };
  
  private handleSearchChange = (value: string) => {
    this.setState({ searchQuery: value });
  };

  private filteredCustomers = () => {
    const { searchQuery, customers } = this.state;
    return customers.filter((customer) =>
      customer.name.toLowerCase().indexOf(searchQuery.toLowerCase()) !== -1
    );
  };

  private addMobileNumber = () => {
    this.setState((prevState) => ({
      newCustomer: {
        ...prevState.newCustomer,
        mobileNumbers: [...prevState.newCustomer.mobileNumbers, ''],
      },
    }));
  };

  private handleMobileChange = (index: number, value: string) => {
    const mobileNumbers = [...this.state.newCustomer.mobileNumbers];
    mobileNumbers[index] = value;
    this.setState((prevState) => ({
      newCustomer: { ...prevState.newCustomer, mobileNumbers },
    }));
  };

  public render(): React.ReactElement<ICustomerscreenProps> {
    const { newCustomer, searchQuery, isFormVisible, isDeleteModalVisible } = this.state;

    return (
      <section className={styles.customerscreen}>
        <h2>Customer Management Screen</h2>

        <PrimaryButton className={styles.PrimaryButton} text={isFormVisible ? "Cancel" : "Add Customer"} onClick={this.toggleFormVisibility} />

        {isFormVisible && (
          <div className={styles.formContainer}>
            <TextField label="Customer Name" value={newCustomer.name} onChange={(e, value) => this.handleInputChange('name', value || '')} />
            <TextField label="Customer Address" value={newCustomer.address} onChange={(e, value) => this.handleInputChange('address', value || '')} />
            <TextField label="Customer Email" value={newCustomer.email} onChange={(e, value) => this.handleInputChange('email', value || '')} />
            <TextField label="GST Number" value={newCustomer.gstNumber} onChange={(e, value) => this.handleInputChange('gstNumber', value || '')} />
            <TextField label="Contact Person" value={newCustomer.contactPerson} onChange={(e, value) => this.handleInputChange('contactPerson', value || '')} />

            <div>
              {newCustomer.mobileNumbers.map((mobile, index) => (
                <TextField key={index} label={`Mobile Number ${index + 1}`} value={mobile} onChange={(e, value) => this.handleMobileChange(index, value || '')} />
              ))}
              <PrimaryButton className={styles.PrimaryButton} text="Add Mobile Number" onClick={this.addMobileNumber} />
            </div>
            <PrimaryButton className={styles.SaveDetails} text={this.state.editingIndex === null ? 'Save details' : 'Update Customer'} onClick={this.addOrUpdateCustomer1} />
          </div>
        )}

        <h3>Customer List</h3>
        <TextField placeholder="Search Customers" value={searchQuery} onChange={(e, value) => this.handleSearchChange(value || '')} />

      <div className={styles.customerTableWrapper}>
        <table className={styles.customerTable}>
          <thead>
            <tr>
              <th>Customer Name</th>
              <th>Address</th>
              <th>Email</th>
              <th>GST Number</th>
              <th>Contact Person</th>
              <th>Mobile Numbers</th>
              <th>Actions</th>
            </tr>
          </thead>
          <tbody>
            {this.filteredCustomers().map((customer, index) => (
              <tr key={index}>
                <td>{customer.name}</td>
                <td>{customer.address}</td>
                <td>{customer.email}</td>
                <td>{customer.gstNumber}</td>
                <td>{customer.contactPerson}</td>
                <td>{customer.mobileNumbers.join(', ')}</td>
                <td>
                    <div className={styles.actions}>
                      <button  className={styles.editButton} onClick={() => this.editCustomer(index)}>Edit</button>
                      <button className={styles.deleteButton} onClick={() => this.showDeleteModal(index)}>Delete</button>
                    </div>
                </td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>

        {isDeleteModalVisible && (
          <Modal isOpen={isDeleteModalVisible} onDismiss={this.closeDeleteModal}>
            <div className={styles.modal}>
            <div className={styles.modalContent}>
              <p className={styles.modalMessage}>Are you sure you want to delete this data?</p>
              <div className={styles.buttonGroup}>
                <button onClick={this.deleteCustomer} className={styles.yesButton}>Yes</button>
                <button onClick={this.closeDeleteModal} className={styles.noButton}>No</button>
              </div>
            </div>
          </div>
          </Modal>
        )}
      </section>
    );
  }
}