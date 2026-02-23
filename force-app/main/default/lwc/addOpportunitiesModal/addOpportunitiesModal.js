import LightningModal from "lightning/modal";
import { api } from "lwc";
export default class AddOpportunitiesModal extends LightningModal {
    @api recordId;
}