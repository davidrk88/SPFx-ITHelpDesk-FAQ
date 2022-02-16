import { Version } from '@microsoft/sp-core-library';
import {
	IPropertyPaneConfiguration,
	PropertyPaneTextField,
	PropertyPaneCheckbox,
	PropertyPaneDropdown,
	PropertyPaneToggle
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { escape } from '@microsoft/sp-lodash-subset';
import {
	SPHttpClient,
	SPHttpClientResponse
} from '@microsoft/sp-http';
import {
	Environment,
	EnvironmentType
} from '@microsoft/sp-core-library';

import styles from './FaqPracticeWebPart.module.scss';
import * as strings from 'FaqPracticeWebPartStrings';
import * as JQuery from 'jquery';
import 'jqueryui';
import { SPComponentLoader } from '@microsoft/sp-loader';
// import MyAccordionTemplate from './MyAccordionTemplate';
import MockHttpClient from './MockHttpClient';


export interface IFaqPracticeWebPartProps {
	faqTitle: string;
	description: string;
	targetList: string;
	quesitonFormatFont: string;
	questionFormatColor: string;
	questionFormatSize: string;
}

export interface FAQLists {
	value: FAQList[];
}

export interface FAQList {
	Title: string;
	Answer: string;
	IsActive: string;
	OrderNum: string;
}

export default class FaqPracticeWebPart extends BaseClientSideWebPart <IFaqPracticeWebPartProps> {

	private availableListOptions: any[];

	public constructor() {
		super();

		SPComponentLoader.loadCss('//code.jquery.com/ui/1.11.4/themes/smoothness/jquery-ui.css');
	}

	private _renderListAsync(): void {
		// this._getListSelections();
		// Local Environment
		if (Environment.type === EnvironmentType.Local) {
			this._getMockListData().then((response) => {
				this._renderList(response.value);
			});
		} else if (Environment.type == EnvironmentType.SharePoint || Environment.type == EnvironmentType.ClassicSharePoint) {
			this._getListData().then((response) => {
				this._renderList(response.value);
			});
		}
	}

	private _renderList(items: FAQList[]): void {
		let html: string = '';
		let headerColor: string = '';
		let fontColor: string = this.properties.questionFormatColor;

		switch(this.properties.questionFormatColor) {
			case 'colorRed':
				headerColor = `<h3 class="${styles.qFontColorRed}">`;
				break;
			case 'colorGreen':
				headerColor = `<h3 class="${styles.qFontColorGreen}">`;
				break;
			case 'colorBlue':
				headerColor = `<h3 class="${styles.qFontColorBlue}">`;
				break;
			default:
				headerColor = `<h3 class="${styles.qFontColorBlack}">`;
		}

		items.forEach((item: FAQList) => {
			// let el = new DOMParser().parseFromString(item.Answer, "text/html");
			// usage - el.documentElement.textContent
			// html += `
			// 	<h3 class="${styles.qFontColorRed}">${ item.Title }</h3>
			// 	<div>
			// 		<p>${ item.Answer }</p>
			// 	</div>`;
			html += headerColor + `${ item.Title }</h3>
				<div>
					<p>${ item.Answer }</p>
				</div>`;
		});

		const listContainer: Element = this.domElement.querySelector('#spListContainer');
		listContainer.innerHTML = html;

		const accordionOptions: JQueryUI.AccordionOptions = {
			animate: true,
			collapsible: true,
			icons: {
				header: 'ui-icon-circle-arrow-e',
				activeHeader: 'ui-icon-circle-arrow-s'
			}
		};

		JQuery('.accordion', this.domElement).accordion(accordionOptions);
	}

	private _getListData(): Promise<FAQLists> {
		return this.context.spHttpClient.get(this.context.pageContext.web.absoluteUrl + `/_api/web/lists/getbytitle('${escape(this.properties.targetList)}')/items`, SPHttpClient.configurations.v1)
			.then((response: SPHttpClientResponse) => {
				return response.json();
			});
	}

	private _getMockListData(): Promise<FAQLists> {
		return MockHttpClient.get()
			.then((data: FAQList[]) => {
				var listData: FAQLists = { value: data };
				return listData;
			}) as Promise<FAQLists>;
	}

	private _getAvailableLists(): Promise<any> {
		return this.context.spHttpClient.get(this.context.pageContext.web.absoluteUrl + `/_api/web/lists?$filter=Hidden eq false`, SPHttpClient.configurations.v1)
			.then((response: SPHttpClientResponse) => {
				return response.json();
			});
	}

	private _getListSelections(): Promise<any> {
		let listSelections = [];
		return this._getAvailableLists().then((response) => {
			response.value.map((item) => {
				listSelections.push( { key: item.Title, text: item.Title } );
			});
			return listSelections;
		});
	}

	public render(): void {
		this.domElement.innerHTML = `
			<div class="${ styles.faqPractice }">
				<div class="${ styles.container }">
					<div class="${ styles.row }">
						<div class="${ styles.column }">
							<h2>${escape(this.properties.faqTitle)}</h2>
							<p class="${ styles.subTitle }">${escape(this.properties.description)}</p>
							<a href="https://www.starwars.com/community" class="${ styles.button }">
							<span class="${ styles.label }">Learn more</span>
							</a>
						</div>
					</div>
					<div id="spListContainer" class="accordion" />
				</div>
			</div>`;

		this._renderListAsync();
	}

	protected get dataVersion(): Version {
		return Version.parse('1.0');
	}

	protected onPropertyPaneConfigurationStart(): void {
		this._getListSelections().then((response) => {
			this.availableListOptions = response;
			this.context.propertyPane.refresh();
		});
	}

	protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
		
		console.log("TESTING - ", this.availableListOptions);
		
		return {
			pages: [
			{
				header: {
					description: strings.PropertyPaneDescription
				},
				groups: [
				{
					groupName: strings.BasicGroupName,
					groupFields: [
						PropertyPaneTextField('faqTitle', {
							label: 'FAQ Title'
						}),
						PropertyPaneTextField('description', {
							label: strings.DescriptionFieldLabel
						}),
						PropertyPaneDropdown('targetList', {
							label: 'Source List for FAQ',
							options: this.availableListOptions
						}),
						// PropertyPaneDropdown('quesitonFormatFont', {
						// 	label: 'Set font for question text',
						// 	options: [
						// 		{ key: 'q-font-type-arial', text: 'Arial' },
						// 		{ key: 'q-font-type-Georgia', text: 'Georgia' }
						// 	]
						// }),
						PropertyPaneDropdown('questionFormatColor', {
							label: 'Set color for question text',
							options: [
								{ key: 'colorBlack', text: 'Default' },
								{ key: 'colorRed', text: 'Red' },
								{ key: 'colorGreen', text: 'Green' },
								{ key: 'colorBlue', text: 'Blue' }
							]
						})
						// PropertyPaneTextField('questionFormatSize', {
						// 	label: 'Set size for question text'
						// })
					]
				}]
			}]
		};
	}
}
