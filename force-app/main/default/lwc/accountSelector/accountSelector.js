// Import the LightningElement base class from the LWC (Lightning Web Components) framework
import { LightningElement, api } from 'lwc';
import { NavigationMixin } from 'lightning/navigation';

export default class AccountSelector extends NavigationMixin(LightningElement) {
    @api selectedAccount;
    isLoading = true; // Initially, set to true to show spinner

    // Method to handle form load event and manage spinner visibility
    handleLoad() {
        // Hide spinner because the form has finished loading 
        this.isLoading = false;

        // Find the Account lookup field using the data-id attribute
        const lookupField = this.template.querySelector('lightning-input-field[data-id="accountLookup"]');

        // Attempt to set focus to the Account lookup field
        lookupField.focus();
    }

    // Method to handle changes in account change
    handleAccountChange = (event) => {
        // Retrieve the selected Account ID from the event's detail
        let accountId = event.detail.value?.toString();
        this.selectedAccount = accountId;
        // Dispatch the custom event to notify the parent component of the account Change
        this.dispatchEvent(new CustomEvent('accountchange', {
            detail: { accountId }  // Include the Account ID in the event's detail object
        }));
    }

    handleViewHierarchy() {
        const componentDef = {
            componentDef: "sfa:hierarchyFullView",
            attributes: {
                recordId: this.selectedAccount,
                sObjectName: "Account"
            }
        };

        const encoded = encodeURIComponent(btoa(JSON.stringify(componentDef)));

        const fullUrl = `/one/one.app#${encoded}`;
        window.open(fullUrl, '_blank');
    }
}