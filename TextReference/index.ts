import { IInputs, IOutputs } from "./generated/ManifestTypes";
import { Dirent } from "fs";
import { encode } from "punycode";
import { EntityDefinition } from "./models/EntityDefinition";
import { env } from "process";
import { Record } from "./models/Record";
import { SSL_OP_MICROSOFT_BIG_SSLV3_BUFFER } from "constants";

export class TextReference implements ComponentFramework.StandardControl<IInputs, IOutputs> {

	private _notifyOutputChanged: () => void;

	private _context: ComponentFramework.Context<IInputs>;
	private _container: HTMLDivElement;
	private _labels: HTMLDivElement;
	private _currentInput: HTMLInputElement | undefined;
	private _records: HTMLUListElement;
	private _position: Range | undefined;

	private _entityDefinitionExclamation: EntityDefinition[];
	private _entityDefinitionAt: EntityDefinition[];
	private _entityDefinitionHashtag: EntityDefinition[];
	private _entityDefinitionDollar: EntityDefinition[];

	/**
	 * Empty constructor.
	 */
	constructor() {
		this._entityDefinitionExclamation = new Array();
		this._entityDefinitionAt = new Array();
		this._entityDefinitionHashtag = new Array();
		this._entityDefinitionDollar = new Array();
	}

	/**
	 * Used to initialize the control instance. Controls can kick off remote server calls and other initialization actions here.
	 * Data-set values are not initialized here, use updateView.
	 * @param context The entire property bag available to control via Context Object; It contains values as set up by the customizer mapped to property names defined in the manifest, as well as utility functions.
	 * @param notifyOutputChanged A callback method to alert the framework that the control has new outputs ready to be retrieved asynchronously.
	 * @param state A piece of data that persists in one session for a single user. Can be set at any point in a controls life cycle by calling 'setControlState' in the Mode interface.
	 * @param container If a control is marked control-type='standard', it will receive an empty div element within which it can render its content.
	 */
	public init(context: ComponentFramework.Context<IInputs>, notifyOutputChanged: () => void, state: ComponentFramework.Dictionary, container: HTMLDivElement) {

		this._notifyOutputChanged = notifyOutputChanged;

		this._context = context;

		this.EntityDefinitions();
		container.append(this._labels);

		this.EntitySelector();
		container.append(this._records);

		this.EditableGrid();
		container.append(this._container);

		if (this._context.parameters.text.raw) {
			this._container.innerHTML = this._context.parameters.text.raw;
			this.AddLinks();
		}
	}

	//MAIN COMPONENTS -----------------------------------------------------------------------------------------------------------------------------------
	private EntityDefinitions() {
		this._labels = document.createElement("div");
		this._labels.id = "text-reference-labels";

		if (this._context.parameters.exclamation.raw) {
			this._entityDefinitionExclamation = this.SplitEntities(this._context.parameters.exclamation.raw!);
			this.AddLabel("!", this._entityDefinitionExclamation);
		}

		if (this._context.parameters.at.raw) {
			this._entityDefinitionAt = this.SplitEntities(this._context.parameters.at.raw!);
			this.AddLabel("@", this._entityDefinitionAt);
		}

		if (this._context.parameters.hashtag.raw) {
			this._entityDefinitionHashtag = this.SplitEntities(this._context.parameters.hashtag.raw!);
			this.AddLabel("#", this._entityDefinitionHashtag);
		}

		if (this._context.parameters.dollar.raw) {
			this._entityDefinitionDollar = this.SplitEntities(this._context.parameters.dollar.raw!);
			this.AddLabel("$", this._entityDefinitionDollar);
		}
	}
	private EntitySelector() {
		this._records = document.createElement("ul");
		this._records.id = "text-reference-records-list";
	}
	private EditableGrid() {
		this._container = document.createElement("div");
		this._container.id = "text-reference-editable-div";
		this._container.tabIndex = 0;
		this._container.dir = "ltr";
		this._container.contentEditable = this._context.mode.isControlDisabled ? "false" : "true";
		this._container.setAttribute("captureKey", "true");
		this._container.addEventListener("mouseup", this.SaveSelection.bind(this));
		this._container.addEventListener("keyup", this.SaveSelection.bind(this));
		this._container.addEventListener("focus", this.SaveSelection.bind(this));
		this._container.addEventListener("keypress", this.CaptureKey.bind(this));
	}

	//DAO -----------------------------------------------------------------------------------------------------------------------------------------------
	private RetrieveEntityMetada(entityLogicalName: string): EntityDefinition | undefined {

		let entityDefinition: EntityDefinition | undefined;
		entityDefinition = undefined;
		let req = new XMLHttpRequest();
		const baseUrl = (<any>this._context).page.getClientUrl();
		const caller = this;
		req.open("GET", baseUrl + "/api/data/v9.1/EntityDefinitions(LogicalName='" + entityLogicalName + "')?$select=IconSmallName,ObjectTypeCode,DisplayName,EntitySetName,PrimaryNameAttribute,PrimaryIdAttribute,EntityColor", false);
		req.setRequestHeader("OData-MaxVersion", "4.0");
		req.setRequestHeader("OData-Version", "4.0");
		req.setRequestHeader("Accept", "application/json");
		req.setRequestHeader("Content-Type", "application/json; charset=utf-8");
		req.onreadystatechange = function () {
			if (this.readyState === 4) {
				req.onreadystatechange = null;
				if (this.status === 200) {
					var result = JSON.parse(this.response);
					entityDefinition = new EntityDefinition();
					entityDefinition.LogicalName = entityLogicalName;
					entityDefinition.EntitySetName = result.EntitySetName;
					entityDefinition.PrimaryIdAttribute = result.PrimaryIdAttribute;
					entityDefinition.PrimaryNameAttribute = result.PrimaryNameAttribute;
					entityDefinition.UserLocalizedLabel = result.DisplayName.UserLocalizedLabel;
					if (result.ObjectTypeCode >= 10000 && result.IconSmallName != null)
						entityDefinition.Icon = baseUrl + "/WebResources/" + result.IconSmallName.toString();
					else
						entityDefinition.Icon = baseUrl + caller.GetURL(result.ObjectTypeCode);
				}
			}
		};
		req.send();

		return entityDefinition;
	}
	private GetURL(objectTypeCode: number) {

		//default icon
		var url = "/_imgs/svg_" + objectTypeCode.toString() + ".svg";

		if (!this.UrlExists(url)) {
			url = "/_imgs/ico_16_" + objectTypeCode.toString() + ".gif";

			if (!this.UrlExists(url)) {
				url = "/_imgs/ico_16_"
					+ objectTypeCode.toString() +
					".png";

				//default icon

				if (!this.UrlExists(url)) {
					url = "/_imgs/ico_16_customEntity.gif";
				}
			}
		}

		return url;
	}
	private UrlExists(url: string) {
		var http = new XMLHttpRequest();
		http.open('HEAD', url, false);
		http.send();
		return http.status != 404 && http.status != 500;
	}
	private ExecuteQuickFind(entityNames: string, value: string): Record[] {

		let records: Record[];
		records = new Array<Record>();

		let splitEnities = entityNames.split(';');
		if (splitEnities && splitEnities.length > 0) {

			let body: any;
			body = new Object();
			body.EntityGroupName = null;
			body.EntityNames = new Array();
			for (let index = 0; index < splitEnities.length; index++) {
				const entityName = splitEnities[index];
				body.EntityNames.push(entityName);
			}
			body.SearchText = value;

			const caller = this;
			var req = new XMLHttpRequest();
			req.open("POST", (<any>this._context).page.getClientUrl() + "/api/data/v9.0/ExecuteQuickFind", false);
			req.setRequestHeader("OData-MaxVersion", "4.0");
			req.setRequestHeader("OData-Version", "4.0");
			req.setRequestHeader("Accept", "application/json");
			req.setRequestHeader("Content-Type", "application/json; charset=utf-8");
			req.onreadystatechange = function () {
				if (this.readyState === 4) {
					req.onreadystatechange = null;
					if (this.status === 200) {
						var response: any;
						response = JSON.parse(this.response);
						if (response && response.Result.length > 0) {
							for (let i = 0; i < response.Result.length; i++) {
								const entity_ = response.Result[i].Data;
								if (entity_ && entity_.length > 0) {
									for (let y = 0; y < entity_.length; y++) {
										const record_ = entity_[y];
										let entityLogicalName = record_["@odata.type"]!.split('.')[3];
										let entityDefinition = caller.GetEntityDefinition(entityLogicalName);

										let record = new Record();
										record.Id = record_[entityDefinition!.PrimaryIdAttribute];
										record.Name = record_[entityDefinition!.PrimaryNameAttribute];
										record.LogicalName = entityLogicalName;

										records.push(record);
									}
								}
							}
						}
					}
				}
			};
			req.send(JSON.stringify(body));
		}

		return records;
	}
	private SplitEntities(parameter: string): EntityDefinition[] {
		let entityDefinitions: EntityDefinition[];
		entityDefinitions = new Array<EntityDefinition>();
		if (parameter) {
			let split = parameter.split(';');
			for (let index = 0; index < split.length; index++) {
				const entityLogicalName = split[index];
				if (entityLogicalName) {
					let entityDefinition = this.RetrieveEntityMetada(entityLogicalName);
					if (entityDefinition) {
						entityDefinitions.push(entityDefinition);
					}
				}
			}
		}

		return entityDefinitions;
	}

	//INTERFACE -----------------------------------------------------------------------------------------------------------------------------------------
	private AddLabel(char: string, entityDefinitions: EntityDefinition[]) {
		let label = document.createElement("label");
		label.innerText = char;
		for (let index = 0; index < entityDefinitions.length; index++) {
			const element = entityDefinitions[index];
			let image = document.createElement("img");
			image.src = element.Icon;
			this._labels.append(image);
			label.innerText = label.innerText + " " + element.UserLocalizedLabel.Label;
			label.style.color = this.GetEntityHexadecimal(element.LogicalName);
		}
		this._labels.append(label);
	}
	private SaveSelection() {
		this._position = window.getSelection()?.getRangeAt(0);
		this.Persist();
	}
	private CaptureKey(ke: KeyboardEvent) {

		if (this._container.getAttribute("captureKey") == "false")
			return;

		switch (ke.keyCode) {
			//!
			case 33:
				if (this._entityDefinitionExclamation.length > 0) {
					ke.preventDefault();
					this.EntitySearch(this._context.parameters.exclamation.raw!);
				}
				break;

			//@
			case 64:
				if (this._entityDefinitionAt.length > 0) {
					ke.preventDefault();
					this.EntitySearch(this._context.parameters.at.raw!);
				}
				break;

			//#
			case 35:
				if (this._entityDefinitionHashtag.length > 0) {
					ke.preventDefault();
					this.EntitySearch(this._context.parameters.hashtag.raw!);
				}
				break;

			//$
			case 36:
				if (this._entityDefinitionDollar.length > 0) {
					ke.preventDefault();
					this.EntitySearch(this._context.parameters.dollar.raw!);
				}
				break;
		}

		this.Persist();
	}
	private EntitySearch(logicalName: string): void {
		let entitySeatchInput: HTMLInputElement;
		entitySeatchInput = document.createElement("input");
		entitySeatchInput.id = "00000000-0000-0000-0000-000000000000";
		entitySeatchInput.setAttribute("type", "text");
		entitySeatchInput.setAttribute("logicalName", logicalName);
		entitySeatchInput.setAttribute("class", "entity-selector");
		entitySeatchInput.addEventListener("keyup", this.EntityFilterKeyPress.bind(this));
		entitySeatchInput.addEventListener("focusout", this.EntityFilterLostFocus.bind(this, entitySeatchInput));
		entitySeatchInput.style.backgroundColor = "whitesmoke";
		this._position!.insertNode(entitySeatchInput);
		entitySeatchInput.focus();
	}
	private EntityFilterKeyPress(ke: KeyboardEvent) {

		this._container.setAttribute("captureKey", "false");

		this._currentInput = ke.currentTarget as HTMLInputElement;
		switch (ke.keyCode) {
			//Backspace
			case 8:
				if (!this._currentInput.value) {
					this._currentInput.remove();
					this._container.focus();
				}
				break;

			//Enter
			case 13:
				ke.preventDefault();
				if (this._currentInput.value && this._currentInput.value.length >= 3) {
					this.ClearRecords();
					let records = this.ExecuteQuickFind(this._currentInput.getAttribute("logicalName")!, this._currentInput.value);

					if (records.length == 1) {
						this._currentInput.id = records[0].Id;
						this.CreateButton(records[0].Id, records[0].LogicalName, records[0].Name);
					}

					else if (records.length > 1) {
						for (let index = 0; index < records.length; index++) {
							const record_ = records[index];

							let record = document.createElement("li");
							record.id = record_.Id;
							record.setAttribute("logicalName", record_.LogicalName);
							record.innerText = record_.Name;
							record.addEventListener("click", this.SelectRecord.bind(this, record));

							let recordImg = document.createElement("img");
							recordImg.src = this.GetEntityDefinition(record_.LogicalName)!.Icon;
							record.append(recordImg);

							this._records.append(record);
						}
						this._records.style.display = "block";
						this._currentInput.after(this._records);
						this._records.focus();
					}
				}
				break;
		}

		this.Persist();
	}
	private EntityFilterLostFocus(caller: HTMLInputElement) {
		this._container.setAttribute("captureKey", "true");
	}
	private SelectRecord(caller: HTMLLIElement) {
		this.CreateButton(caller.id, caller.getAttribute("logicalName")!, caller.innerText);
	}
	private CreateButton(id: string, logicalName: string, name: string) {

		let button = document.createElement("div");
		button.id = id;
		button.setAttribute("class", "entity-record");
		button.setAttribute("logicalName", logicalName);
		button.style.color = this.GetEntityHexadecimal(logicalName);
		button.style.width = ((name.length + 1) * 9) + 'px';
		button.contentEditable = "false";
		this.AddLink(button);

		let img = document.createElement("img");
		img.src = this.GetEntityDefinition(logicalName)!.Icon;
		button.append(img);

		let lbl = document.createElement("label");
		lbl.innerText = name;
		button.append(lbl);

		this._currentInput!.replaceWith(button);
		this._currentInput = undefined;
		this._records.style.display = "none";

		this.Persist();
	}
	private AddLinks() {
		let divs = document.getElementsByClassName("entity-record");
		if (divs) {
			for (let index = 0; index < divs.length; index++) {
				const element = divs[index] as HTMLDivElement;
				this.AddLink(element);
			}
		}
	}
	private AddLink(button: HTMLDivElement) {
		button.addEventListener("click", this.OpenForm.bind(this, button));
	}
	private OpenForm(div: HTMLDivElement) {

		var parameters: any;
		parameters = {};
		parameters.entityId = div.id;
		parameters.entityName = div.getAttribute("logicalName");
		parameters.openInNewWindow = true;
		parameters.windowPosition = 1;
		this._context.navigation.openForm(parameters);
	}
	private ClearRecords() {
		if (this._records) {
			while (this._records.firstChild) {
				this._records.removeChild(this._records.firstChild);
			}
		}
		this._records.style.display = "none";
	}

	//LOCAL DATA -----------------------------------------------------------------------------------------------------------------------------------------
	private GetEntityDefinition(entityLogicalName: string): EntityDefinition | undefined {
		let entityDefinition: EntityDefinition | undefined;
		if (this._entityDefinitionExclamation && this._entityDefinitionExclamation.length > 0) {
			entityDefinition = this._entityDefinitionExclamation.find(f => f.LogicalName == entityLogicalName);
			if (entityDefinition)
				return entityDefinition;
		}
		if (this._entityDefinitionHashtag && this._entityDefinitionHashtag.length > 0) {
			entityDefinition = this._entityDefinitionHashtag.find(f => f.LogicalName == entityLogicalName);
			if (entityDefinition)
				return entityDefinition;
		}
		if (this._entityDefinitionDollar && this._entityDefinitionDollar.length > 0) {
			entityDefinition = this._entityDefinitionDollar.find(f => f.LogicalName == entityLogicalName);
			if (entityDefinition)
				return entityDefinition;
		}
		return entityDefinition;
	}
	private GetEntityHexadecimal(entityLogicalName: string): string {
		if (this._entityDefinitionExclamation.filter(f => f.LogicalName == entityLogicalName).length > 0)
			return this._context.parameters.exclamation_color.raw!;
		else if (this._entityDefinitionAt.filter(f => f.LogicalName == entityLogicalName).length > 0)
			return this._context.parameters.at_color.raw!;
		else if (this._entityDefinitionHashtag.filter(f => f.LogicalName == entityLogicalName).length > 0)
			return this._context.parameters.hashtag_color.raw!;
		else if (this._entityDefinitionDollar.filter(f => f.LogicalName == entityLogicalName).length > 0)
			return this._context.parameters.dollar_color.raw!;
		else
			return "black";
	}
	public Persist() {
		this._notifyOutputChanged();
	}

	/**
	 * Called when any value in the property bag has changed. This includes field values, data-sets, global values such as container height and width, offline status, control metadata values such as label, visible, etc.
	 * @param context The entire property bag available to control via Context Object; It contains values as set up by the customizer mapped to names defined in the manifest, as well as utility functions
	 */
	public updateView(context: ComponentFramework.Context<IInputs>): void {
		this._context = context;
	}

	/** 
	 * It is called by the framework prior to a control receiving new data. 
	 * @returns an object based on nomenclature defined in manifest, expecting object[s] for property marked as “bound” or “output”
	 */
	public getOutputs(): IOutputs {
		return { text: this._container.innerHTML };
	}

	/** 
	 * Called when the control is to be removed from the DOM tree. Controls should use this call for cleanup.
	 * i.e. cancelling any pending remote calls, removing listeners, etc.
	 */
	public destroy(): void {
		// Add code to cleanup control if necessary
	}
}