/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global console, document, Excel, Office */

// The initialize function must be run each time a new page is loaded
Office.onReady(() => {
	const sideload = document.getElementById("sideload-msg");
	if (sideload) sideload.style.display = "none";
	const app = document.getElementById("app-body");
	if (app) app.style.display = "flex";

	const insertBtn = document.getElementById("insert-names-btn");
	if (insertBtn) (insertBtn as HTMLButtonElement).onclick = insertNames;

	const readBtn = document.getElementById("read-names-btn");
	if (readBtn) (readBtn as HTMLButtonElement).onclick = readNames;
});

export async function insertNames() {
	try {
		await Excel.run(async (context: Excel.RequestContext) => {
			const sheet = context.workbook.worksheets.getActiveWorksheet();
			sheet.names.add("udf_name", "=ADD(1,2)");
			sheet.names.add("sum_name", "=SUM(1,2)");
			await context.sync();
		});
	} catch (error) {
		console.error(error);
	}
}

function formatNamedItem(item: Excel.NamedItem): string {
	return (
		`Name:    ${item.name}\n` +
		`Type:    ${item.type}\n` +
		`Formula: ${item.formula}\n` +
		`Value:   ${JSON.stringify(item.value)}`
	);
}

export async function readNames() {
	const resultEl = document.getElementById("read-result");
	if (resultEl) resultEl.textContent = "Reading…";
	try {
		await Excel.run(async (context: Excel.RequestContext) => {
			const sheet = context.workbook.worksheets.getActiveWorksheet();
			const udfItem = sheet.names.getItem("udf_name");
			const sumItem = sheet.names.getItem("sum_name");
			udfItem.load(["name", "value", "type", "formula"]);
			sumItem.load(["name", "value", "type", "formula"]);
			await context.sync();
			const text =
				`--- udf_name ---\n${formatNamedItem(udfItem)}\n\n` +
				`--- sum_name ---\n${formatNamedItem(sumItem)}`;
			if (resultEl) resultEl.textContent = text;
		});
	} catch (error) {
		if (resultEl) resultEl.textContent = `Error: ${error}`;
		console.error(error);
	}
}
