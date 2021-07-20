import { IInputs, IOutputs } from "./generated/ManifestTypes";
import "./js/jquery.js";
import { CdsService } from "./viewmodels/CdsService";
import { SharePointService } from "./viewmodels/SharePointService";
import $ = require("jquery");



export class LargeFileUpload implements ComponentFramework.StandardControl<IInputs, IOutputs> {

	//input values of fields will be stored here
	private _sharePointSiteGUID: string | null;
	private _sharePointStructureEntity: string | null;
	private _clientID: string | null;
	private _sharePointRelativeURL: string | null;
	private _sharePointAboluteUrl: string | null;
	private _loginHint: string;
	private _folder: string | null;

	//PCF context containing the parameters, control metadata and interface functions.
	private _context: ComponentFramework.Context<IInputs>;

	//Control's container is an Div Element
	private _container: HTMLDivElement;

	//input element created for this control
	private _Uploadinput: HTMLInputElement;

	// PCF delegate which will be assigned to this object which would be called when any update happens. 
	private _notifyOutputChanged: () => void;

	//Button element created for this control (can be replace by input type submit)
	private uploadButton: HTMLButtonElement;

	//Create service instance
	private serviceSP: SharePointService;

	//File variable
	private file: File | undefined | null;


	/**
	 * Empty constructor.
	 */
	constructor() {

	}
	public init(context: ComponentFramework.Context<IInputs>, notifyOutputChanged: () => void, state: ComponentFramework.Dictionary, container: HTMLDivElement): void {
		// Add control initialization code
		this._context = context;

		this._notifyOutputChanged = notifyOutputChanged;

		//this._value = context.parameters.sampleProperty.raw;
		this._clientID = context.parameters.clientID.raw;
		this._sharePointStructureEntity = context.parameters.sharePointStructureEntity.raw;
		this._folder = context.parameters.folder.raw;
		this._sharePointSiteGUID = context.parameters.sharePointSiteGUID.raw;
		this._sharePointRelativeURL = context.parameters.sharePointRelativeURL.raw;
		this._loginHint = "dev1@msftnbu.com";

		//Create an upload button to upload file
		this._Uploadinput = document.createElement("input");
		this._Uploadinput.id = "FileUploader";
		this._Uploadinput.name = "FileUploader";
		this._Uploadinput.value = "Select File";
		this._Uploadinput.type = "file"
		this._Uploadinput.accept = ".*";
		//Added a few styles settings can move this into css file
		this._Uploadinput.style.opacity = "1";
		this._Uploadinput.style.width = "auto";
		this._Uploadinput.style.height = "auto";
		this._Uploadinput.style.pointerEvents = "all";

		//create submit button
		this.uploadButton = document.createElement("button");
		this.uploadButton.id = "SubmitButton";
		this.uploadButton.name = "SubmitButton";
		this.uploadButton.value = "Submit";
		this.uploadButton.innerText = "Submit";
		this.uploadButton.className = "button"



		//Add event Listeners
		this._Uploadinput.addEventListener("change", this.GetMyFile.bind(this));
		this._Uploadinput.addEventListener("click", this.GetMyFile.bind(this));
		this.uploadButton.addEventListener("click", this.onSubmitButtonClick.bind(this));

		//Adding elements created to the DIV container
		this._container = document.createElement("div");
		this._container.appendChild(this._Uploadinput);
		this._container.appendChild(this.uploadButton);

		//Display container
		container.appendChild(this._container);

	}


	/**
	 * Called when any value in the property bag has changed. This includes field values, data-sets, global values such as container height and width, offline status, control metadata values such as label, visible, etc.
	 * @param context The entire property bag available to control via Context Object; It contains values as set up by the customizer mapped to names defined in the manifest, as well as utility functions
	 */
	public updateView(context: ComponentFramework.Context<IInputs>): void {
		this._context = context;
	}

	public getOutputs(): IOutputs {
		return {};
	}

	public destroy(): void {
	}

	//Implement code to submit file into server
	private onSubmitButtonClick(): void {
		this.serviceSP =
		new SharePointService(this._context, {
			sharePointStructureEntity: this._context.parameters.sharePointStructureEntity.raw ?? "",
			sharepointSiteId: this._context.parameters.sharePointSiteGUID.raw ?? "",
			clientId: this._context.parameters.clientID.raw ?? "",
			loginHint: this._context.parameters.LoginHint.raw ?? "",
		}

			)
		try {
			this.serviceSP.setupSharePoint(<string>this._loginHint, <string>this._sharePointAboluteUrl);
			this.serviceSP.uploadFileToSharePoint(<string>this._folder, this.file);
		}
		catch (e) {
			console.log((e as Error).message);
		}

	}


	//This function loads the selected file
	private GetMyFile(event: Event): void {

		this.file = this._Uploadinput.files?.item(0);

		//var FileElement: Document
		//$("#fileUploader").change(function (event) {
		//	var FileElement: <HTMLInputElement>document.querySelector('#FileUploader');
		//	var file: any;
		//	var files: any;

		//	if (FileElement != null) {
		//		files = FileElement.files;
		//		if (files?.length != null) {
		//			file = files[0];
		//		}
		//	}

			
		//	//Implement reader
		//	reader.onerror = function (event) {
		//		console.error("File could not be load! Code ");
		//	};

		//	reader.readAsBinaryString(<Blob>file);
		//});

		this._notifyOutputChanged();

		
	}
}