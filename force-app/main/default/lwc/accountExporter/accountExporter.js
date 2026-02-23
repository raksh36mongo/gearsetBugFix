/**
 * @description Lightning Web Component (LWC) to generate a spreadsheet displaying a consolidated view of an Account's Renewal Opportunities.
 *
 * Features:
 * - Loads XLSX library dynamically for spreadsheet generation.
 * - Fetches renewal opportunity data from Apex controller based on Account recordId.
 * - Formats and styles spreadsheet data, including headers, cells, hyperlinks, and borders.
 * - Handles both open and closed (prior) renewal opportunities, including related Opportunity Line Items and Quotes.
 * - Provides utility methods for currency formatting, cell styling, and border application.
 * - Displays toast notifications for success, error, and informational messages.
 * - Triggers file download of the generated spreadsheet and closes the action screen upon completion.
 *
 * @date 4 August 2025
 *
 * @class
 * @extends LightningElement
 *
 * @property {boolean} loading - Indicates if the component is in a loading state.
 * @property {string} _recordId - Internal storage for the Account record Id.
 * @property {Set<string>} OLIfieldAPISet - Set of field API names for Opportunity Line Items.
 * @property {Set<string>} OppfieldAPISet - Set of field API names for open Opportunities.
 * @property {Set<string>} priorOppFieldAPISet - Set of field API names for prior Opportunities.
 * @property {Object} priorOppFieldMap - Field label mapping for prior Opportunities.
 * @property {Object} openOppFieldMap - Field label mapping for open Opportunities.
 * @property {Array} OLIHeaders - Styled header cells for Opportunity Line Items.
 * @property {Array} oppHeaders - Styled header cells for Opportunities.
 *
 * @method renderedCallback - Loads the XLSX library when the component is rendered.
 * @method styleHeaders - Styles header cells with borders and font.
 * @method styleCell - Styles individual cells with font and fill.
 * @method parseParentObjectField - Retrieves parent field values from child records.
 * @method connectedCallback - Sets loading state when component is initialized.
 * @method fetchConfigData - Fetches configuration and opportunity data from Apex, sets headers and field sets.
 * @method formatDataForSpreadsheet - Formats fetched data for spreadsheet generation.
 * @method createHyperlink - Creates a styled hyperlink cell for spreadsheet.
 * @method formatOpportunityWithRelatedRecord - Formats opportunity data with related records for spreadsheet rows.
 * @method formatCurrency - Formats numeric values as currency strings.
 * @method applyBorders - Applies outer borders to a 2D data array for spreadsheet.
 * @method formatRenewalOpportunity - Formats a renewal opportunity and its prior contracts into a 2D array.
 * @method handleSuccessToast - Displays a success toast notification.
 * @method handleErrorToast - Displays an error toast notification.
 * @method handleInfoToast - Displays an informational toast notification.
 * @method generateExcel - Generates and triggers download of the spreadsheet file.
 *
 * @fires ShowToastEvent - For displaying toast notifications.
 * @fires CloseActionScreenEvent - For closing the action screen after completion.
 */
/**
 * @description LWC to generate an Spreadsheet to display Account's Renewal Opportunities consolidated view.
 * @Date 4 August 2025
 */

import { LightningElement, api, track } from 'lwc'
import { loadScript } from 'lightning/platformResourceLoader'
import xlsxbundle from '@salesforce/resourceUrl/xlsxbundle'
import { CloseActionScreenEvent } from 'lightning/actions'
import populateAccountOpportunitiesData from '@salesforce/apex/AccountOpportunityDataController.populateAccountOpportunitiesData'
import { ShowToastEvent } from 'lightning/platformShowToastEvent'
import OpportunityName from '@salesforce/schema/Opportunity.Name'
import OpportunityId from '@salesforce/schema/Opportunity.Id'
import QuoteName from "@salesforce/schema/SBQQ__Quote__c.Name";
import QuoteID from '@salesforce/schema/SBQQ__Quote__c.Id'
const blankValues = ['N/A' , false , undefined , null ,'']

export default class AccountExporter extends LightningElement {
    loading = true

    _recordId
    @api set recordId(value) {
        // set record Id and get data from Apex
        this._recordId = value
        this.fetchConfigData()
    }

    get recordId() {
        return this._recordId
    }

    // callback method to load the XLSX Library
    async renderedCallback() {
        try {
            await loadScript(this, xlsxbundle + '/xlsxbundle/xlsx.bundle.js')
        } catch (err) {
            console.log({ err })
        }
    }

    // utility method to style cells and set correct type
    styleCell(cell, fontSize, boldFont, bgColour , wrapCell) {
        if (cell !== undefined && typeof cell === "string" && (cell.startsWith("http") || cell.startsWith("https"))) {
            return this.createHyperlink(cell, undefined, boldFont, bgColour);
        }

        const style = {
            font: { bold: boldFont || false, sz: fontSize || 8 }
        };
        if (bgColour) {
            style.fill = { fgColor: { rgb: bgColour } };
        }

        if(wrapCell){
               style.alignment = {wrapText: true}
               // 3. Add thin borders to all cells
      style.border = {
          top: { style: "thin", color: { rgb: "FFD3D3D3" } },
          left: { style: "thin", color: { rgb: "FFD3D3D3" } },
          bottom: { style: "thin", color: { rgb: "FFD3D3D3" } },
          right: { style: "thin", color: { rgb: "FFD3D3D3" } }
      };

        }

        // Ensure a valid value always exists
        let value = ''
        const regex = /^[A-Z]{3}\s\d{1,3}(?:,\d{3})*(?:\.\d{1,2})?$/
        if (cell !== undefined && cell !== null) {
            if (typeof cell === 'string' && cell.length > 3 && regex.test(cell)) {
                const getCurrencySymbol = cell.substring(0, 3)
                value = cell.replace(getCurrencySymbol, '')
            } else {
                value = cell
            }
        }
        // Set the correct Excel type: 'n' for number, 's' for string
        const type = typeof value === "number" ? "n" : "s";
        // If it’s numeric, apply formatting with commas + 2 decimals
        

        return { v: value, t: type, s: style , z : "#,##0.00"};
    }

    // utility method to parse parent field value from the child record(eg. Opportunity->Account.Name)
    parseParentObjectField(record, fieldAPI) {
        console.log('record, fieldAPI', record, fieldAPI)
        if (record != undefined) {
            return record[fieldAPI]
        }
        return ''
    }

    connectedCallback() {
        this.loading = true
    }

    // method to fetch configuration data from Apex and set the headers and field API sets
    async fetchConfigData() {
        try {
            const response = await populateAccountOpportunitiesData({ accId: this.recordId }) // get renewal opportunities data from Apex
            console.log('response', response)
            //  const jsObj = JSON.parse(JSON.parse(JSON.stringify(response)))
            if (response == undefined) {
                this.handleInfoToast('Error while fetching data.Please contact your system administrator.')
                this.loading = false
                this.dispatchEvent(new CloseActionScreenEvent())
                return
            }

            // parse the response to get the renewal opportunities data
            const jsObj = JSON.parse(response)
            if (jsObj.renewalOpportunities == undefined || jsObj.renewalOpportunities?.length === 0) {
                this.handleInfoToast('No Open Renewal Opportunities found for this Account.')
                this.loading = false
                this.dispatchEvent(new CloseActionScreenEvent())
                return
            }

            console.log('jsObj', jsObj)

            this.OLIHeaders = Object.values(jsObj.OLIHeaderConfig)
                .reverse()
                .map(element => {
                    return this.styleCell(element, 10, true)
                })

            this.oppHeaders = Object.values(jsObj.oppHeaderConfig)
                .reverse()
                .map(element => {
                    return this.styleCell(element, 10, true)
                })

            this.OLIfieldAPISet = new Set(Object.keys(jsObj.OLIHeaderConfig).reverse())
            this.OppfieldAPISet = new Set(Object.keys(jsObj.oppHeaderConfig).reverse())
            this.priorOppFieldAPISet = new Set(Object.keys(jsObj.priorOppHeaderConfig).reverse())

            this.priorOppFieldMap = jsObj.priorOppHeaderConfig
            this.openOppFieldMap = jsObj.oppHeaderConfig

            this.formatDataForSpreadsheet(jsObj.renewalOpportunities, jsObj.accInfo.Name) // call method to format data for the spreadsheet
        } catch (err) {
            console.log({ err })
            this.handleErrorToast(err);

            this.loading = false;
            this.dispatchEvent(new CloseActionScreenEvent());
            return;
        }
    }

    // method to generate data from the current Account record. This data would be used in the spreadsheet later.
    async formatDataForSpreadsheet(renewalOpportunitiesArr, accountName) {
        try {
            //  const AccountName = jsObj.accInfo.Name

            let renewalData = []

            renewalOpportunitiesArr?.forEach((renewalOpp, index) => {
                const renewalOppDataArr = this.formatRenewalOpportunity(renewalOpp) // call method to format renewal opportunity data | This will return a 2D array

                renewalData.push(...renewalOppDataArr)
            })

            this.generateExcel({ data: renewalData, AccountName: accountName })
        } catch (err) {
            console.log('err', err)
        }
    }

    // utility method to create hyperlink on the cells
    createHyperlink(linkToDisplay, recordId, boldFont, bgColour) {
        // utility method to style cells
        const style = {
            font: { bold: boldFont || false, sz: 9, color: { rgb: '0563C1' }, underline: true } // Set font color to blue and underline for hyperlink
        }
        if (bgColour) {
            style.fill = { fgColor: { rgb: bgColour } }
        }

        return {
            v: linkToDisplay,
            l:
                recordId != undefined
                    ? { Target: `https://${window.location.hostname}/${recordId}` }
                    : { Target: linkToDisplay },
            t: 's', // Ensure it's treated as a string
            s: style
        }
    }

    // method to format opportunity with products
    formatOpportunityWithRelatedRecord(opportunityRecord, oppFieldMap, priorOppCount) {
        let primaryQuoteHeaderObj = {
            Approval_Justification__c: 'Commercial Terms',
            SBQQ__PaymentTerms__c: 'Payment Terms',
            Cancellation_Policy__c: 'Evergreen/Marketplace Extended Term Contract Language',
            Additional_Terms__c: 'Additional Terms Language',
            Consulting_Terms__c: 'Consulting Terms',
            T_E_Cap__c: 'T&E Terms',
            Basic_Customer_Reference__c: 'Basic Customer Reference',
            Customer_References__c: 'Strategic Customer Reference',
            Remove_Auto_Renewal_Language__c: 'Auto-Renewal Terms',
            Add_Price_Cap_Additional__c: 'Pricing Caps',
            Customer_Agreement_Legal__c: 'Agreements and T&C Language'
        }
        try {
            let renewalOppArr = []

            // add opportunity rows

            const opportunityData = opportunityRecord.opportunityData || opportunityRecord

            if (!opportunityRecord.IsClosed) {
                // to add header when opportunity is open
                renewalOppArr.push([this.styleCell('Open Renewal Opportunity', 10, true, '9dd7ef')])
            } else {
                // to add header when opportunity is closed
                renewalOppArr.push([this.styleCell(`Prior Support Contract - ${priorOppCount}`, 9, true, 'c6ccd3')])
            }

            renewalOppArr.push([
                this.createHyperlink(
                    opportunityData[OpportunityName.fieldApiName],
                    opportunityData[OpportunityId.fieldApiName],
                    true,
                    opportunityData.IsClosed ? 'c6ccd3' : '9dd7ef'
                )
            ])
            renewalOppArr.push([this.styleCell()], [this.styleCell()]) // adding spaces on rows

            renewalOppArr.push(this.OLIHeaders) // adding Opportunity Line Item headers

            const OLIData =
                opportunityRecord?.opportunityData?.OpportunityLineItems?.records ||
                opportunityRecord?.OpportunityLineItems?.records // get opportunity products
            if (OLIData?.length > 0) {
                // if there are products in the opportunity
                OLIData.forEach(product => {
                    // iterate over product objects
                    let OLIRows = []

                    this.OLIfieldAPISet.forEach(fieldAPI => {
                        if (fieldAPI.includes('__r.')) {
                            // to check if field is a parent object field
                            OLIRows.push(this.styleCell(this.parseParentObjectField(product, fieldAPI), 9))
                        } else {
                            OLIRows.push(this.styleCell(this.formatCurrency(product[fieldAPI]), 9))
                        }
                    })

                    renewalOppArr.push(OLIRows)
                })
            } // Opportunity Line Items are added on the sheet

            renewalOppArr.push([this.styleCell()], [this.styleCell()]) // adding spaces on rows

            renewalOppArr.push([this.styleCell('Opportunity Data', 10, true)]) // adding header for opportunity data header

            for (let oppField in opportunityData) {
                // looping over opportunity fields
                if (oppField.includes('__r')) {
                    // check if field is a lookup field
                    // if lookup field
                    renewalOppArr.push([this.styleCell()], [this.styleCell()])
                    for (let innerField in opportunityData[oppField]) {
                        if (primaryQuoteHeaderObj[innerField] != undefined) {
                            renewalOppArr.push([this.styleCell(primaryQuoteHeaderObj[innerField], 10, true)])
                        }

                        if (innerField === QuoteName.fieldApiName) {
                            renewalOppArr.push([
                                this.styleCell("Primary Quote", 10, true),
                                this.createHyperlink(
                                    opportunityData[oppField][innerField],
                                    opportunityData[oppField][QuoteID.fieldApiName],
                                    false,
                                    ""
                                )
                            ]);
                        } else if ((
                            innerField == "Reseller_Agreement_Legal__c" ||
                            innerField == "Cloud_Subscription_Agreement_Legal__c" ||
                            innerField == "Customer_Agreement_Legal__c") &&  !blankValues.includes(opportunityData[oppField][innerField])
                        ) {
                            renewalOppArr.push([
                                this.styleCell(oppFieldMap[oppField + "." + innerField], 9),
                                this.createHyperlink("view", opportunityData[oppField][innerField], false, "")
                            ]);
                        } else if (
                            this.priorOppFieldAPISet.has(oppField + "." + innerField) &&
                            !blankValues.includes(this.parseParentObjectField(opportunityData[oppField], innerField))
                        ) {
                            renewalOppArr.push([
                                this.styleCell(oppFieldMap[oppField + "." + innerField], 9),
                                this.styleCell(this.parseParentObjectField(opportunityData[oppField], innerField), 9 , false , false , true)
                            ]);
                        }
                    }
                } else {
                    // opportunity original fields
                    if (oppField != OpportunityName.fieldApiName && oppFieldMap[oppField] != undefined) {
                        renewalOppArr.push([
                            this.styleCell(oppFieldMap[oppField], 9, !opportunityRecord.IsClosed),
                            this.styleCell(
                                this.formatCurrency(opportunityData[oppField]),
                                9,
                                !opportunityRecord.IsClosed
                            )
                        ])
                    }

                    if (oppField === 'Subscription_Months__c' && !opportunityRecord.IsClosed) {
                        // to add previous contract opportunities in the middle oof Open Opportunity data
                        opportunityRecord?.previousContractOpp?.forEach((opp, index) => {
                            renewalOppArr.push([
                                this.styleCell(`Previous Support Contract - ${index + 1} `, 9, true),
                                this.createHyperlink(opp[OpportunityName.fieldApiName], opp[OpportunityId.fieldApiName])
                            ])
                        })
                    }
                }
            }

            this.applyBorders(renewalOppArr) // apply borders to the data array(outer border)

            renewalOppArr.push([this.styleCell('')], [this.styleCell('')]) // adding spaces on rows

            return renewalOppArr
        } catch (err) {
            console.log({ err })
        }
    }

    // utility method to format currency values
    formatCurrency(value) {
        if (typeof value == 'number') {
            if (Number.isInteger(value)) {
                // Just format with commas, no decimals
                return value.toString().replace(/\B(?=(\d{3})+(?!\d))/g, ',')
            } else {
                // Keep 2 decimals for non-integers
                return value.toFixed(2).replace(/\B(?=(\d{3})+(?!\d))/g, ',')
            }
        } else {
            return !isNaN(Number(value.replace(/,/g, ''))) && value.trim() !== ''
                ? Number(value.replace(/,/g, ''))
                    .toFixed(2)
                    .replace(/\B(?=(\d{3})+(?!\d))/g, ',')
                : value
        }
    }

    // utility method to apploy border to whole Opportunity data array(individual Opportunity table)
    applyBorders(dataArr) {
        const columnLength = Math.max(...dataArr.map(a => a.length))

        // Apply only to outer cells
        dataArr.forEach((row, rowIndex) => {
            for (let column = 0; column < columnLength; column++) {
                if (!row[column]) {
                    row[column] = { v: '', s: {} }
                }
                const borderStyle = getBorder(rowIndex, column, dataArr, columnLength)
                if (Object.keys(borderStyle).length > 0) {
                    row[column].s = row[column].s || {}
                    row[column].s.border = borderStyle
                }
            }
        })

        // method to get border style for each cell
        function getBorder(row, column, dataArr, columnLength) {
            const border = {}
            const black = { style: 'thin', color: { rgb: 'FF000000' } } // SheetJS border format

            const lastRowIndex = dataArr.length - 1
            const lastColIndex = columnLength - 1

            // Top border for first row
            if (row === 0) border.top = black

            // Bottom border for last row
            if (row === lastRowIndex) border.bottom = black

            // Left border for first column
            if (column === 0) border.left = black

            // Right border for last column
            if (column === lastColIndex) border.right = black

            return border
        }
    }

    // method to format renewal opportunity
    formatRenewalOpportunity(renewalOpp) {
        let renewalOpp2DArray = []

        // Add Open Renewal Opportunity record
        const openRenewalOpportunityArr = this.formatOpportunityWithRelatedRecord(renewalOpp, this.openOppFieldMap)
        if (openRenewalOpportunityArr != undefined && openRenewalOpportunityArr?.length != 0) {
            renewalOpp2DArray.push(...openRenewalOpportunityArr)
        }

        renewalOpp2DArray.push([], [])

        //handle Previous opportunity contract(closed)

        renewalOpp.previousContractOpp.forEach((previousOpportunity, index) => {
            const previousOppData = this.formatOpportunityWithRelatedRecord(
                previousOpportunity,
                this.priorOppFieldMap,
                index + 1
            )
            renewalOpp2DArray.push(...previousOppData)
        })

        return renewalOpp2DArray
    }

    // util method to display success toast
    handleSuccessToast() {
        this.dispatchEvent(
            new ShowToastEvent({
                title: 'Success',
                message: 'Data fetched successfully. Please check your download folder.',
                variant: 'success'
            })
        )
    }

    // util method to display error toast
    handleErrorToast(error) {
        this.dispatchEvent(
            new ShowToastEvent({
                title: "Error",
                message: error.body
                    ? error.body.message
                    : "An error occurred while fetching data.Please contact your system administrator.",
                variant: "error"
            })
        );
    }

    // util method to display error toast
    handleInfoToast(message) {
        this.dispatchEvent(
            new ShowToastEvent({
                title: 'Unable to fetch data',
                message: message,
                variant: 'info'
            })
        )
    }

    // method to generate and download the spreadsheet with the Account's data
    async generateExcel(formattedData) {
        try {
            const sheetData = formattedData.data

            // Create worksheet and workbook
            const worksheet = XLSX.utils.aoa_to_sheet(sheetData)

            // Auto width with minimum fallback
            const maxCols = Math.max(...sheetData.map(row => row.length))
            worksheet['!cols'] = Array.from({ length: maxCols }).map((_, i) => {
                const maxLen = Math.max(
                    ...sheetData.map(row => {
                        const cell = row[i]
                        const val = typeof cell === 'object' ? cell?.v : cell
                        return val ? val.toString().length : 0
                    })
                )

                let width = Math.max(12, maxLen + 2) // default min = 12 chars
                return { wch: width }
            })

            const workbook = XLSX.utils.book_new()
            XLSX.utils.book_append_sheet(workbook, worksheet, 'Renewals')
            // Force Column B (index 1) width
            if (!worksheet['!cols']) worksheet['!cols'] = []
            worksheet['!cols'][1] = { wch: 40 } // fixed width = 30 chars
            // Trigger download
            XLSX.writeFile(workbook, `${formattedData.AccountName} Renewal Opportunities Data.xlsx`)

            this.handleSuccessToast()
        } catch (err) {
            console.log('err', err)
            this.handleErrorToast(err)
        } finally {
            this.loading = false
            this.dispatchEvent(new CloseActionScreenEvent())
        }
    }
}