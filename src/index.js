const zip = require('@zip.js/zip.js');

const nameWithPathList = [
	{
		name: 'code',
		path: '//w:body/w:p[1]/w:r[2]/w:t',
	},
];

const inputElement = document.getElementById('file');
inputElement.onchange = async function (event) {
	const file = inputElement.files[0];

	const reader = new zip.ZipReader(new zip.BlobReader(file));
	const entries = await reader.getEntries();

	const xmlText = await getXML(entries);

	const parser = new DOMParser();
	const XMLdocument = parser.parseFromString(xmlText, 'application/xml');

	console.log('XMLdocument: ', XMLdocument);

	const values = getValues(XMLdocument, nameWithPathList, nsResolver);

	printResult(values);
};

async function getXML(entries) {
	for (let key in entries) {
		const entry = entries[key];
		if (entry.filename === 'word/document.xml') {
			xmlText = await entry.getData(new zip.TextWriter());
			return xmlText;
		}
	}
}

function printResult(values) {
	values.forEach((item) => console.log(item.name + ' = ' + item.value));
}

function getValues(xml, valuePathList, nsResolver) {
	const res = valuePathList.map((valuePath) => {
		if (xml.evaluate) {
			var nodes = xml.evaluate(valuePath.path, xml, nsResolver, XPathResult.ANY_TYPE, null);
			var result = nodes.iterateNext();
			const value = result.childNodes[0].nodeValue;
			return {
				name: valuePath.name,
				value,
			};
		}
	});
	return res;
}

function nsResolver(prefix) {
	var ns = {
		w: 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
	};
	return ns[prefix] || null;
}
