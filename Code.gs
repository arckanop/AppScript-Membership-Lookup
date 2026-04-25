function myFunction() {
	// if (MailApp.getRemainingDailyQuota() == 0.0) return;
	// else Logger.log(MailApp.getRemainingDailyQuota());
  
	// Logger.log(MailApp.getRemainingDailyQuota());
	// retryEmailByRow(167);

	// retryEmailByRow(865);
	// <!-- Stop -->

	// for (let i = 826; i <= 832; i++) {
	// 	retryEmailByRow(i);
	// }

	// let rows = [834, 835, 836, 837, 838, 841, 845, 849, 850, 851, 858, 859, 860, 862, 877, 878, 880, 881, 888, 891, 892, 893, 895, 897 , 898];

	// for (let row of rows) {
	// 	retryEmailByRow(row);
	// }
}

function doGet() {
	getData();
	return HtmlService
        .createHtmlOutputFromFile('index.html')
        .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function getData() {
	const cache = CacheService.getScriptCache();
	const cached = cache.get("examMap");
	if (cached) return JSON.parse(cached);

	const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Form Responses");
	const values = sheet.getDataRange().getValues();

	const examMap = Object.create(null);

	for (let i = 1; i < values.length; i++) {
		const examId = String(values[i][9]).trim();
		if (!examId) continue;

		examMap[examId] = {
			id: String(values[i][2] ?? "").trim(),
			phone: String(values[i][7] ?? "").trim(),
		};
	}

	try {
		cache.put("examMap", JSON.stringify(examMap), 600);
	} catch (e) {
		Logger.log("Cache put failed: " + e.message);
	}

	return examMap;
}

function findID(examId, phone) {
	try {
		const examMap = getData();
		const key = String(examId).trim();

		if (!Object.prototype.hasOwnProperty.call(examMap, key)) {
			return { found: false };
		}

		const entry = examMap[key];
		if (String(entry.phone).trim() !== String(phone).trim()) {
			return { found: false };
		}

		return { found: true, membershipId: entry.id };
	} catch (error) {
		return { found: false, error: error.message };
	}
}

function invalidateExamCache() {
	CacheService.getScriptCache().remove("examMap");
}

function retryEmailByRow(row) {
	const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Form Responses");
	if (!sheet) throw new Error('Response sheet not found');

	row = Number(row);
	if (!row || row < 2) throw new Error('Invalid row number');

	try {
		if (MailApp.getRemainingDailyQuota() < 1) throw new Error('No email quota remaining today');

		const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
		const values = sheet.getRange(row, 1, 1, sheet.getLastColumn()).getValues()[0];

		const record = {};
		headers.forEach((h, i) => record[h] = values[i]);

		const membershipID = String(values[2] || '').trim();
		if (!membershipID) throw new Error('No membership ID in column C');

		const emailAddress = String(record['Email Address'] || record['ที่อยู่อีเมล'] || '').trim();
		if (!emailAddress) throw new Error('Missing email address');

		const emailData = {
			membershipID: membershipID,
			parentFirstName: record["ชื่อผู้ปกครอง"] || '',
			parentLastName: record["นามสกุลผู้ปกครอง"] || '',
			studentFirstName: record["ชื่อนักเรียน"] || '',
			studentLastName: record["นามสกุลนักเรียน"] || '',
			jerseySize: record["ไซส์เสื้อ Jersey"] || '-',
			jerseyText: record["ตัวอักษรบนเสื้อ Jersey"] || '-',
			jerseyNumber: record["ตัวเลขบนเสื้อ Jersey"] || '-',
			poloSize: record["ไซส์เสื้อโปโล"] || '-',
			poloGender: record["เพศเสื้อโปโล"] || '-'
		};

		const subject = 'ขอบคุณสำหรับการกรอกแบบฟอร์มสมัครสมาชิกและจองไซส์เสื้อ Triamudom Family';
		const plainText = generatePlainText(emailData);
		const htmlBody = generateHtmlBody(emailData);

		GmailApp.sendEmail(emailAddress, subject, plainText, {
			htmlBody: htmlBody,
			name: 'ข้อมูลการสมัครสมาชิก',
			noReply: true
		});

		sheet.getRange(row, 32).setValue("TRUE");

		sheet
			.getRange(row, 1, 1, sheet.getLastColumn())
			.setFontFamily('Anuphan');

		const monoRanges = [
			'A' + row, 'B' + row, 'C' + row, 'H' + row, 'J' + row,
			'Q' + row, 'R' + row, 'S' + row, 'T' + row, 'AD' + row
		];
		sheet.getRangeList(monoRanges).setFontFamily('JetBrains Mono');

		sheet.getRange('A' + row).setHorizontalAlignment('right');
		sheet.getRangeList(['B' + row, 'AD' + row]).setHorizontalAlignment('left');
		sheet.getRangeList([
			'C' + row, 'H' + row, 'I' + row, 'J' + row, 'K' + row,
			'L' + row, 'M' + row, 'Q' + row, 'R' + row, 'S' + row,
			'T' + row, 'U' + row, 'AE' + row
		]).setHorizontalAlignment('center');

		sheet.getRange(row, 33).setValue('Fixed');
		console.log('Retry succeeded for row %s', row);
	} catch (err) {
		sheet.getRange(row, 33).setValue('Retry Failed: ' + err.message);
		throw err;
	}
}

function onFormSubmit(e) {
	const lock = LockService.getDocumentLock() || LockService.getScriptLock();
	lock.waitLock(30000);

	try {
		const formData = e.namedValues;
		const responseSheet = e.range.getSheet();
		const referenceSheet = responseSheet.getParent().getSheetByName('CurrentNumber');

		if (!referenceSheet) throw new Error('Sheet "CurrentNumber" not found');

		const track = (formData['แผนการเรียน']?.[0] || '').trim();

		if (!track) throw new Error("Track not specified in the form response");

		const counterCell = referenceSheet.getRange(3, getColumn(track));
		const membershipNumber = (Number(counterCell.getValue()) || 0) + 1;
		counterCell.setValue(membershipNumber);

		const membershipID = "89-" + String(membershipNumber);

		const row = e.range.getRow();
		responseSheet.getRange(row, 3).setValue(membershipID);

		invalidateExamCache()

		const emailAddress = formData['Email Address']?.[0] || formData['ที่อยู่อีเมล']?.[0] || '';
		if (emailAddress) {
			const subject = 'ขอบคุณสำหรับการกรอกแบบฟอร์มสมัครสมาชิกและจองไซส์เสื้อ Triamudom Family';

			const emailData = {
				membershipID: membershipID,
				parentFirstName: formData["ชื่อผู้ปกครอง"]?.[0] || '',
				parentLastName: formData["นามสกุลผู้ปกครอง"]?.[0] || '',
				studentFirstName: formData["ชื่อนักเรียน"]?.[0] || '',
				studentLastName: formData["นามสกุลนักเรียน"]?.[0] || '',
				jerseySize: formData["ไซส์เสื้อ Jersey"]?.[0] || '-',
				jerseyText: formData["ตัวอักษรบนเสื้อ Jersey"]?.[0] || '-',
				jerseyNumber: formData["ตัวเลขบนเสื้อ Jersey"]?.[0] || '-',
				poloSize: formData["ไซส์เสื้อโปโล"]?.[0] || '-',
				poloGender: formData["เพศเสื้อโปโล"]?.[0] || '-'
			};

			const plainText = generatePlainText(emailData);
			const htmlBody = generateHtmlBody(emailData);

			GmailApp.sendEmail(emailAddress, subject, plainText, {
				htmlBody: htmlBody,
				name: 'ข้อมูลการสมัครสมาชิก',
				noReply: true
			});

			responseSheet.getRange(row, 32).setValue("TRUE");
		}

		responseSheet
			.getRange(row, 1, 1, responseSheet.getLastColumn())
			.setFontFamily('Anuphan');

		const monoRanges = ['A' + row, 'B' + row, 'C' + row, 'H' + row, 'J' + row, 'Q' + row, 'R' + row, 'S' + row, 'T' + row, 'AD' + row];
		responseSheet.getRangeList(monoRanges).setFontFamily('JetBrains Mono');

		responseSheet.getRange('A' + row).setHorizontalAlignment('right');
		responseSheet.getRangeList(['B' + row, 'AD' + row]).setHorizontalAlignment('left');
		responseSheet.getRangeList(['C' + row, 'H' + row, 'I' + row, 'J' + row, 'K' + row, 'L' + row, 'M' + row, 'Q' + row, 'R' + row, 'S' + row, 'T' + row, 'U' + row, 'AE' + row]).setHorizontalAlignment('center');
	} finally {
		lock.releaseLock();
	}
}

function getColumn(track) {
	switch (track) {
		case 'ภาษา - ภาษาฝรั่งเศส':
			return 2;
		case 'ภาษา - ภาษาเยอรมัน':
			return 3;
		case 'ภาษา - ภาษาญี่ปุ่น':
			return 4;
		case 'ภาษา - ภาษาจีน':
			return 5;
		case 'ภาษา - ภาษาสเปน':
			return 6;
		case 'ภาษา - ภาษาเกาหลี':
			return 7;
		case 'ภาษา - คณิตศาสตร์':
			return 8;
		case 'วิทยาศาสตร์ - คณิตศาสตร์':
			return 9;
		default:
			throw new Error("Invalid Track: " + track);
	}
}

function generatePlainText(data) {
	return `
ขอบคุณสำหรับการสมัครสมาชิก

เรียนคุณ ${data.parentFirstName} ${data.parentLastName}

ขอบคุณสำหรับการกรอกแบบฟอร์มสมัครสมาชิก Triamudom Family และจองไซส์เสื้อ
ทางเราได้รับข้อมูลของท่านเรียบร้อยแล้ว

หมายเลขสมาชิก
${data.membershipID}

ข้อมูลสมาชิก
- ชื่อผู้ปกครอง: ${data.parentFirstName} ${data.parentLastName}
- ชื่อนักเรียน: ${data.studentFirstName} ${data.studentLastName}

ข้อมูลเสื้อ Jersey
- ไซส์เสื้อ Jersey: ${data.jerseySize}
- ชื่อบนเสื้อ Jersey: ${data.jerseyText}
- เบอร์บนเสื้อ Jersey: ${data.jerseyNumber}

ข้อมูลเสื้อโปโล
- ไซส์เสื้อโปโล: ${data.poloSize}
- เพศเสื้อโปโล: ${data.poloGender}

หากมีข้อมูลเพิ่มเติม ทางทีมงานจะแจ้งให้ท่านทราบอีกครั้ง

ขอขอบพระคุณอีกครั้ง

อีเมลฉบับนี้ถูกส่งโดยอัตโนมัติ กรุณาอย่าตอบกลับอีเมลนี้
	`.trim();
}

function generateHtmlBody(data) {
	return `
		<div
			style="
				background-color: #f3f4f6;
				padding: 24px;
				font-family:
					&quot;Anuphan&quot;, &quot;Noto Sans Thai&quot;, &quot;Kanit&quot;,
					&quot;IBM Plex Sans Thai&quot;, Tahoma, Arial, sans-serif;
			"
		>
			<div
				style="
					max-width: 768px;
					margin: 0 auto;
					overflow: hidden;
					border: 1px solid #e5e7eb;
					border-radius: 16px;
					background-color: #ffffff;
					box-shadow:
						0 1px 3px rgba(0, 0, 0, 0.1),
						0 1px 2px rgba(0, 0, 0, 0.06);
				"
			>
				<div
					style="
						padding: 36px 40px;
						background: linear-gradient(to bottom left, #f472b6, #f9a8d4);
						color: #ffffff;
					"
				>
					<h1
						style="margin: 0; font-size: 30px; line-height: 36px; font-weight: 700"
					>
						ขอบคุณสำหรับการสมัครสมาชิก
					</h1>
					<p
						style="
							margin: 8px 0 0 0;
							font-size: 14px;
							line-height: 20px;
							color: #fdf2f8;
						"
					>
						เราได้รับข้อมูลของท่านเรียบร้อยแล้ว
					</p>
				</div>

				<div style="padding: 40px; color: #1f2937">
					<p
						style="
							margin: 0 0 16px 0;
							font-size: 16px;
							line-height: 24px;
							font-weight: 700;
						"
					>
						เรียนคุณ ${escapeHtml_(data.parentFirstName)}
						${escapeHtml_(data.parentLastName)}
					</p>

					<p
						style="
							margin: 0 0 24px 0;
							font-size: 14px;
							line-height: 28px;
							color: #374151;
						"
					>
						ขอบคุณสำหรับการกรอกแบบฟอร์มสมัครสมาชิก Triamudom Family และจองไซส์เสื้อ
					</p>

					<div
						style="
							margin: 0 0 32px 0;
							padding: 28px 24px;
							text-align: center;
							border: 2px solid #fbcfe8;
							border-radius: 16px;
							background-color: #fdf2f8;
						"
					>
						<p
							style="
								margin: 0;
								font-size: 14px;
								line-height: 20px;
								font-weight: 600;
								letter-spacing: 0.1em;
								text-transform: uppercase;
								color: #ec4899;
							"
						>
							Membership Number
						</p>
						<p
							style="
								margin: 8px 0 0 0;
								font-size: 48px;
								line-height: 48px;
								font-weight: 800;
								letter-spacing: 0.025em;
								color: #db2777;
								font-family:
									&quot;JetBrains Mono&quot;, &quot;Fira Code&quot;,
									&quot;Geist Mono&quot;, &quot;Noto Sans Mono&quot;,
									&quot;Source Code Pro&quot;, &quot;Martian Mono&quot;,
									&quot;IBM Plex Mono&quot;, monospace;
							"
						>
							${escapeHtml_(String(data.membershipID))}
						</p>
					</div>

					<div
						style="
							margin: 0 0 32px 0;
							padding: 24px;
							border: 1px solid #e5e7eb;
							border-radius: 12px;
							background-color: #f9fafb;
						"
					>
						<div
							style="
								margin: 0 0 12px 0;
								font-size: 12px;
								line-height: 16px;
								font-weight: 700;
								letter-spacing: 0.05em;
								text-transform: uppercase;
								color: #6b7280;
							"
						>
							ข้อมูลสมาชิก
						</div>
						<p
							style="margin: 0; font-size: 14px; line-height: 20px; color: #1f2937"
						>
							<strong>ชื่อผู้ปกครอง:</strong> ${escapeHtml_(data.parentFirstName)}
							${escapeHtml_(data.parentLastName)}
						</p>
						<p
							style="
								margin: 8px 0 0 0;
								font-size: 14px;
								line-height: 20px;
								color: #1f2937;
							"
						>
							<strong>ชื่อนักเรียน:</strong> ${escapeHtml_(data.studentFirstName)}
							${escapeHtml_(data.studentLastName)}
						</p>
					</div>

					<table
						role="presentation"
						cellpadding="0"
						cellspacing="0"
						border="0"
						width="100%"
						style="margin: 0 0 32px 0; border-collapse: separate; border-spacing: 0"
					>
						<tr>
							<td valign="top" width="50%" style="padding-right: 12px">
								<div
									style="
										border: 1px solid #e5e7eb;
										border-radius: 16px;
										overflow: hidden;
										background-color: #ffffff;
									"
								>
									<div
										style="
											padding: 16px 20px;
											border-bottom: 1px solid #e5e7eb;
											background-color: #fdf2f8;
										"
									>
										<h3
											style="
												margin: 0;
												font-size: 16px;
												line-height: 24px;
												font-weight: 700;
												color: #1f2937;
											"
										>
											ข้อมูลเสื้อโปโล
										</h3>
									</div>

									<div style="padding: 16px 20px; border-bottom: 1px solid #e5e7eb">
										<table
											role="presentation"
											cellpadding="0"
											cellspacing="0"
											border="0"
											width="100%"
										>
											<tr>
												<td
													style="
														font-size: 14px;
														line-height: 20px;
														font-weight: 500;
														color: #6b7280;
													"
												>
													ไซส์เสื้อโปโล
												</td>
												<td
													align="right"
													style="
														font-size: 14px;
														line-height: 20px;
														font-weight: 600;
														color: #1f2937;
													"
												>
													${escapeHtml_(data.poloSize)}
												</td>
											</tr>
										</table>
									</div>

									<div style="padding: 16px 20px">
										<table
											role="presentation"
											cellpadding="0"
											cellspacing="0"
											border="0"
											width="100%"
										>
											<tr>
												<td
													style="
														font-size: 14px;
														line-height: 20px;
														font-weight: 500;
														color: #6b7280;
													"
												>
													เพศเสื้อโปโล
												</td>
												<td
													align="right"
													style="
														font-size: 14px;
														line-height: 20px;
														font-weight: 600;
														color: #1f2937;
													"
												>
													${escapeHtml_(data.poloGender)}
												</td>
											</tr>
										</table>
									</div>
								</div>
							</td>

							<td valign="top" width="50%" style="padding-left: 12px">
								<div
									style="
										border: 1px solid #e5e7eb;
										border-radius: 16px;
										overflow: hidden;
										background-color: #ffffff;
									"
								>
									<div
										style="
											padding: 16px 20px;
											border-bottom: 1px solid #e5e7eb;
											background-color: #fdf2f8;
										"
									>
										<h3
											style="
												margin: 0;
												font-size: 16px;
												line-height: 24px;
												font-weight: 700;
												color: #1f2937;
											"
										>
											ข้อมูลเสื้อ Jersey
										</h3>
									</div>

									<div style="padding: 16px 20px; border-bottom: 1px solid #e5e7eb">
										<table
											role="presentation"
											cellpadding="0"
											cellspacing="0"
											border="0"
											width="100%"
										>
											<tr>
												<td
													style="
														font-size: 14px;
														line-height: 20px;
														font-weight: 500;
														color: #6b7280;
													"
												>
													ไซส์เสื้อ Jersey
												</td>
												<td
													align="right"
													style="
														font-size: 14px;
														line-height: 20px;
														font-weight: 600;
														color: #1f2937;
													"
												>
													${escapeHtml_(data.jerseySize)}
												</td>
											</tr>
										</table>
									</div>

									<div style="padding: 16px 20px; border-bottom: 1px solid #e5e7eb">
										<table
											role="presentation"
											cellpadding="0"
											cellspacing="0"
											border="0"
											width="100%"
										>
											<tr>
												<td
													style="
														font-size: 14px;
														line-height: 20px;
														font-weight: 500;
														color: #6b7280;
													"
												>
													ชื่อบนเสื้อ Jersey
												</td>
												<td
													align="right"
													style="
														font-size: 14px;
														line-height: 20px;
														font-weight: 600;
														color: #1f2937;
													"
												>
													${escapeHtml_(data.jerseyText)}
												</td>
											</tr>
										</table>
									</div>

									<div style="padding: 16px 20px">
										<table
											role="presentation"
											cellpadding="0"
											cellspacing="0"
											border="0"
											width="100%"
										>
											<tr>
												<td
													style="
														font-size: 14px;
														line-height: 20px;
														font-weight: 500;
														color: #6b7280;
													"
												>
													เบอร์บนเสื้อ Jersey
												</td>
												<td
													align="right"
													style="
														font-size: 14px;
														line-height: 20px;
														font-weight: 600;
														color: #1f2937;
													"
												>
													${escapeHtml_(data.jerseyNumber)}
												</td>
											</tr>
										</table>
									</div>
								</div>
							</td>
						</tr>
					</table>

					<p
						style="
							margin: 0 0 24px 0;
							font-size: 14px;
							line-height: 28px;
							color: #374151;
						"
					>
						หากมีข้อมูลเพิ่มเติม ทางทีมงานจะแจ้งให้ท่านทราบอีกครั้ง
					</p>

					<p style="margin: 0; font-size: 14px; line-height: 20px; color: #1f2937">
						ขอขอบคุณอีกครั้ง<br />
						<span style="font-weight: 600">ทีมงาน Triamudom Family</span>
					</p>

					<p
						style="
							margin: 32px 0 0 0;
							font-size: 12px;
							line-height: 16px;
							color: #6b7280;
						"
					>
						อีเมลฉบับนี้ถูกส่งโดยอัตโนมัติ กรุณาอย่าตอบกลับอีเมลนี้
					</p>
				</div>

				<div
					style="
						padding: 16px 40px;
						border-top: 1px solid #e5e7eb;
						background-color: #f9fafb;
						text-align: center;
					"
				>
					<span style="font-size: 14px; line-height: 20px; color: #4b5563">
						<a
							href="https://github.com/arckanop/AppScript-ID-Lookup/tree/membership-lookup"
							target="_blank"
							rel="noopener noreferrer"
							style="
								color: #4b5563;
								text-decoration: none;
								display: inline-block;
								vertical-align: middle;
							"
						>
							<svg
								xmlns="http://www.w3.org/2000/svg"
								viewBox="0 0 24 24"
								fill="#4b5563"
								width="20"
								height="20"
								style="
									display: inline-block;
									vertical-align: middle;
									margin-right: 6px;
								"
							>
								<path
									d="M12 2C6.477 2 2 6.59 2 12.253c0 4.53 2.865 8.37 6.839 9.727.5.095.682-.222.682-.494 0-.244-.009-.89-.014-1.747-2.782.617-3.37-1.37-3.37-1.37-.454-1.178-1.11-1.492-1.11-1.492-.908-.636.069-.623.069-.623 1.004.072 1.532 1.06 1.532 1.06.892 1.574 2.341 1.119 2.91.856.091-.664.35-1.119.636-1.376-2.221-.259-4.555-1.137-4.555-5.062 0-1.118.389-2.032 1.029-2.748-.103-.26-.446-1.302.098-2.714 0 0 .84-.276 2.75 1.05A9.303 9.303 0 0112 6.844c.85.004 1.705.117 2.504.343 1.909-1.326 2.747-1.05 2.747-1.05.546 1.412.203 2.454.1 2.714.64.716 1.028 1.63 1.028 2.748 0 3.936-2.338 4.8-4.566 5.054.359.318.678.946.678 1.907 0 1.377-.012 2.487-.012 2.826 0 .275.18.594.688.493C19.138 20.62 22 16.78 22 12.253 22 6.59 17.523 2 12 2z"
								/>
							</svg>
							<span style="display: inline-block; vertical-align: middle"
								>Made by Arckanop</span
							>
						</a>
						<span style="display: inline-block; vertical-align: middle"> - </span>
						<a
							href="https://github.com/arckanop/AppScript-ID-Lookup/blob/membership-lookup/LICENSE"
							target="_blank"
							rel="noopener noreferrer"
							style="color: #4b5563; text-decoration: none"
						>
							AGPL-3.0 license
						</a>
						<span style="color: #d1d5db"> | </span>
						<span style="display: inline-block; vertical-align: middle"
							>AutoMail v1.2.3</span
						>
					</span>
				</div>
			</div>
		</div>
	`.trim();
}

function escapeHtml_(text) {
	return String(text ?? '')
		.replace(/&/g, '&amp;')
		.replace(/</g, '&lt;')
		.replace(/>/g, '&gt;')
		.replace(/"/g, '&quot;')
		.replace(/'/g, '&#39;');
}