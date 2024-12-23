import * as React from 'react';
import styles from './PoScreen.module.scss';
import { IPoScreenProps } from './IPoScreenProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { Web } from 'sp-pnp-js';

interface IPORecord {
  serialNumber: string;
  rfqNumber: string;
  totalRate: number;
  totalWeight: number;
  totalAmount: number;
  poNumber?: string; // Optional initially until set
  poDate?: string;
  rateByCustomer?: string;
}

interface IPOscreenState {
  serialNumbers: string[];
  selectedSerialNumber: string;
  recordDetails: IPORecord | null;
  showPOForm: boolean;
  poNumber: string;
  poDate: string;
  rateByCustomer: string;
}

export default class PoScreen extends React.Component<IPoScreenProps, IPOscreenState> {
  constructor(props: IPoScreenProps) {
    super(props);

    this.state = {
      serialNumbers: [],
      selectedSerialNumber: '',
      recordDetails: null,
      showPOForm: false,
      poNumber: '',
      poDate: '',
      rateByCustomer: '',
    };
  }

  public componentDidMount(): void {
    this.fetchSerialNumbers();
  }

  private fetchSerialNumbers = async (): Promise<void> => {
    try {
      // Replace with the correct URL for your SharePoint site
     const web = new Web("https://skgroupenginering.sharepoint.com/sites/SalesManagement") // Specify your SharePoint site URL here.
  
      // Fetch RFQ serial numbers where Status is 'Won'
      const items = await web.lists
        .getByTitle('QuotationList')
        .items.filter("Status eq 'Won'") // Filter by Status = 'Won'
        .select('RFQSerialNumber')
        .get();
  
      const serialNumbers = items.map((item: any) => item.RFQSerialNumber);
      this.setState({ serialNumbers });
      console.log("Fetched serial numbers with Status 'Won'");
    } catch (error) {
      console.error('Error fetching serial numbers:', error);
    }
  };
  

  private handleSerialNumberChange = async (event: React.ChangeEvent<HTMLSelectElement>): Promise<void> => {
    const selectedSerialNumber = event.target.value;

    if (!selectedSerialNumber) {
      this.setState({ recordDetails: null, selectedSerialNumber: '' });
      return;
    }

    try {
     const web = new Web("https://skgroupenginering.sharepoint.com/sites/SalesManagement"); // Specify your SharePoint site URL here.

      // Fetch record details from the QuotationList
      const rfqItems = await web.lists
        .getByTitle('QuotationList')
        .items.filter(`RFQSerialNumber eq '${selectedSerialNumber}'`)
        .select('RFQSerialNumber', 'TotalRate', 'TotalWeight', 'TotalAmount')
        .get();

      // Fetch PO details from the POList
      const poItems = await web.lists
        .getByTitle('POList')
        .items.filter(`SerialNumber eq '${selectedSerialNumber}'`)
        .select('PONumber', 'PODate', 'RateByCustomer')
        .get();

      if (rfqItems.length > 0) {
        // Correct field mapping
        const { RFQSerialNumber, TotalRate, TotalWeight, TotalAmount } = rfqItems[0];

        const poDetails = poItems.length > 0 ? poItems[0] : null;

        const recordDetails: IPORecord = {
          serialNumber: selectedSerialNumber,
          rfqNumber: RFQSerialNumber, // Correct field
          totalRate: TotalRate,
          totalWeight: TotalWeight,
          totalAmount: TotalAmount,
          poNumber: poDetails?.PONumber || 'N/A',
          poDate: poDetails?.PODate || 'N/A',
          rateByCustomer: poDetails?.RateByCustomer || 'N/A',
        };

        this.setState({ recordDetails, selectedSerialNumber });
      } else {
        this.setState({ recordDetails: null, selectedSerialNumber });
      }
    } catch (error) {
      console.error('Error fetching record or PO details:', error);
    }
  };


  private openPOForm = (): void => {
    this.setState({ showPOForm: true });
  };

  private closePOForm = (): void => {
    this.setState({
      showPOForm: false,
      poNumber: '',
      poDate: '',
      rateByCustomer: '',
    });
  };

  private handlePOFormChange = (event: React.ChangeEvent<HTMLInputElement>): void => {
    const { name, value } = event.target;
    this.setState({ [name]: value } as any);
  };

  private savePO = async (): Promise<void> => {
    const { poNumber, poDate, rateByCustomer, selectedSerialNumber } = this.state;
  
    if (!poNumber || !poDate || !rateByCustomer) {
      alert('Please fill all fields.');
      return;
    }
  
    try {
      const web = new Web("https://skgroupenginering.sharepoint.com/sites/SalesManagement") // Specify your SharePoint site URL here.
  
      // Check if a record already exists for the selected RFQ number
      const existingPO = await web.lists
        .getByTitle('POList')
        .items.filter(`SerialNumber eq '${selectedSerialNumber}'`)
        .get();
  
      if (existingPO.length > 0) {
        // Update existing record
        await web.lists
          .getByTitle('POList')
          .items.getById(existingPO[0].Id) // Use the ID of the existing record
          .update({
            PONumber: poNumber,
            PODate: poDate,
            RateByCustomer: rateByCustomer,
          });
  
        alert('PO updated successfully.');
      } else {
        // Create a new record
        await web.lists
          .getByTitle('POList')
          .items.add({
            PONumber: poNumber,
            PODate: poDate,
            RateByCustomer: rateByCustomer,
            SerialNumber: selectedSerialNumber,
          });
  
        alert('PO created successfully.');
      }
  
      // Update recordDetails in the state
      this.setState((prevState) => ({
        recordDetails: {
          ...prevState.recordDetails,
          poNumber,
          poDate,
          rateByCustomer,
        } as IPORecord,
      }));
  
      this.closePOForm();
    } catch (error) {
      console.error('Error saving PO:', error);
      alert('Failed to save PO. Please try again.');
    }
  };

  public render(): React.ReactElement<IPoScreenProps> {
    const { serialNumbers, selectedSerialNumber, recordDetails, showPOForm, poNumber, poDate, rateByCustomer } = this.state;

    return (
      <section className={styles.pOscreen}>
        <h2>PO Management Screen</h2>
        <div className={styles.dropdownWrapper}>
        <label 
          htmlFor="rfqSerialNumber" 
          className={styles.dropdownLabel}
        >
          Select RFQ Serial Number:
        </label>
        <select 
          id="rfqSerialNumber" 
          value={selectedSerialNumber} 
          onChange={this.handleSerialNumberChange} 
          className={styles.dropdownSelect}
        >
          <option value="">Select...</option>
          {serialNumbers.map((serial, index) => (
            <option key={index} value={serial}>
              {serial}
            </option>
          ))}
        </select>
      </div>

      {recordDetails && (
      <div className={styles.tableWrapper}>
        <table className={styles.detailsTable}>
          <thead>
            <tr>
              <th>RFQ Number</th>
              <th>Total Rate</th>
              <th>Total Weight</th>
              <th>Total Amount</th>
              <th>PO Number</th>
              <th>PO Date</th>
              <th>Final Rate by Customer</th>
              <th>Action</th>
            </tr>
          </thead>
          <tbody>
            <tr>
              <td>{recordDetails.rfqNumber}</td>
              <td>{recordDetails.totalRate}</td>
              <td>{recordDetails.totalWeight}</td>
              <td>{recordDetails.totalAmount}</td>
              <td>{recordDetails.poNumber}</td>
              <td>{recordDetails.poDate}</td>
              <td>{recordDetails.rateByCustomer}</td>
              <td>
                <button
                  onClick={this.openPOForm}
                  className={`${styles.addButton}`}
                >
                  Add PO
                </button>
              </td>
            </tr>
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
                <input type="text" name="poNumber" value={poNumber} onChange={this.handlePOFormChange} />
              </label>
              <label>
                PO Date:
                <input type="date" name="poDate" value={poDate} onChange={this.handlePOFormChange} />
              </label>
              <label>
                Final Rate by Customer:
                <input type="text" name="rateByCustomer" value={rateByCustomer} onChange={this.handlePOFormChange} />
              </label>
              <div className={styles.formActions}>
                <button onClick={this.savePO} className={`${styles.saveButton}`}>Save</button>
                <button onClick={this.closePOForm} className={`${styles.cancelButton}`}>Cancel</button>
              </div>
            </div>
          </div>
        )}
      </section>
    );
  }
}