import { FAQList } from './FaqPracticeWebPart';

export default class MockHttpClient {
	private static _items: FAQList[] = [
		{ Title: 'Test Q 1', Answer: 'Test ANSWER 1', IsActive: 'Yes', OrderNum: '1' },
		{ Title: 'Test Q 2', Answer: 'Test ANSWER 2', IsActive: 'Yes', OrderNum: '2' },
		{ Title: 'Test Q 3', Answer: 'Test ANSWER 3', IsActive: 'Yes', OrderNum: '3' }
	];

	public static get(): Promise<FAQList[]> {
		return new Promise<FAQList[]>((resolve) => {
			resolve(MockHttpClient._items);
		});
	}
}