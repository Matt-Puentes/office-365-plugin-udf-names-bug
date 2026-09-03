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
	const run2 = document.getElementById("run");
	if (run2) (run2 as HTMLButtonElement).onclick = run;
});

export async function run() {
	try {
		await Excel.run(async (context: Excel.RequestContext) => {
			/**
			 * Insert your Excel code here
			 */
			const range = context.workbook.getSelectedRange();

			// Read the range address
			range.load("address");

			// Update the fill color
			range.format.fill.color = "yellow";

			await context.sync();
			console.log(`The range address was ${range.address}.`);
		});
	} catch (error) {
		console.error(error);
	}
}
