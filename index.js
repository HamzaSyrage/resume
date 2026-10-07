const {
	Document,
	Packer,
	Paragraph,
	TextRun,
	HeadingLevel,
	AlignmentType,
	Numbering,
	LevelFormat,
	BorderStyle,
	ExternalHyperlink,
	convertInchesToTwip,
	UnderlineType,
} = require("docx");
const fs = require("fs");

const path = require("path");
const { execSync } = require("child_process");

const content = require("./content");

const NAVY = "1F2937";
const ACCENT = "334155";
const RULE = "9CA3AF";

const FONT = "Calibri";

const FILE_STEM = "Hamza-Syrage-Frontend-Engineer-Resume";

const hr = () =>
	new Paragraph({
		border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: RULE } },
		spacing: { after: 60 },
	});

const sectionHeading = (text) =>
	new Paragraph({
		spacing: { before: 60, after: 20 },
		border: { bottom: { style: BorderStyle.SINGLE, size: 4, color: RULE } },
		children: [
			new TextRun({
				text: text.toUpperCase(),
				bold: true,
				size: 21,
				color: NAVY,
				font: FONT,
			}),
		],
	});

const jobHeader = (title, dates) =>
	new Paragraph({
		spacing: { before: 40, after: 4 },
		tabStops: [{ type: "right", position: convertInchesToTwip(6.5) }],
		children: [
			new TextRun({
				text: title,
				bold: true,
				size: 22,
				color: NAVY,
				font: FONT,
			}),
			new TextRun({
				text: "\t" + dates,
				italics: true,
				size: 20,
				color: ACCENT,
				font: FONT,
			}),
		],
	});

const subHeader = (text) =>
	new Paragraph({
		spacing: { after: 20 },
		children: [
			new TextRun({ text, italics: true, size: 20, color: ACCENT, font: FONT }),
		],
	});

const bullet = (runsOrText) =>
	new Paragraph({
		numbering: { reference: "bullet-list", level: 0 },
		spacing: { after: 14 },
		children:
			typeof runsOrText === "string"
				? [
						new TextRun({
							text: runsOrText,
							size: 20,
							font: FONT,
							color: "1F2937",
						}),
					]
				: runsOrText,
	});

const bold = (text) =>
	new TextRun({ text, bold: true, size: 20, font: FONT, color: NAVY });
const norm = (text) =>
	new TextRun({ text, size: 20, font: FONT, color: "1F2937" });
const link = (text, url) =>
	new ExternalHyperlink({
		children: [
			new TextRun({
				text,
				style: "Hyperlink",
				size: 18,
				font: FONT,
				color: "334155",
			}),
		],
		link: url.startsWith("http") ? url : `https://${url}`,
	});

const titledLinkParagraph = (title, label, url) =>
	new Paragraph({
		spacing: { before: 40, after: 8 },
		children: [bold(title + " - "), link(label, url)],
	});

const buildChildren = () => {
	const { profile, summary, experience, projects, skills, education } = content;

	const children = [
		// Name
		new Paragraph({
			alignment: AlignmentType.CENTER,
			spacing: { after: 20 },
			children: [
				new TextRun({
					text: profile.name.toUpperCase(),
					bold: true,
					size: 40,
					color: NAVY,
					font: FONT,
				}),
			],
		}),

		// Headline
		new Paragraph({
			alignment: AlignmentType.CENTER,
			spacing: { after: 40 },
			children: [
				new TextRun({
					text: profile.headline,
					size: 22,
					color: ACCENT,
					font: FONT,
				}),
			],
		}),

		// Contact line + profile links
		new Paragraph({
			alignment: AlignmentType.CENTER,
			spacing: { after: 40 },
			children: [
				new TextRun({
					text: `${profile.location}  |  ${profile.phone}  |  ${profile.email}`,
					size: 18,
					font: FONT,
					color: "334155",
				}),
				new TextRun({ text: "", break: 1 }),
				...profile.links.flatMap((l, i) => [
					...(i > 0 ? [norm("  |  ")] : []),
					link(l.label, l.url),
				]),
			],
		}),

		hr(),

		// Summary
		sectionHeading("Professional Summary"),
		new Paragraph({
			spacing: { after: 10 },
			children: [norm(summary)],
		}),
	];

	// Experience
	if (experience.length) {
		children.push(sectionHeading("Professional Experience"));

		for (const job of experience) {
			children.push(
				new Paragraph({
					spacing: { before: 40, after: 4 },
					tabStops: [{ type: "right", position: convertInchesToTwip(6.5) }],
					children: [
						bold(job.role + " - "),
						link(job.company, job.companyUrl),
						new TextRun({
							text: "\t" + job.datesLabel,
							italics: true,
							size: 20,
							color: ACCENT,
							font: FONT,
						}),
					],
				}),
			);

			if (job.note) {
				children.push(
					new Paragraph({
						spacing: { after: 10 },
						children: job.noteUrl
							? [
									new TextRun({
										text: job.note.slice(
											0,
											job.note.length - job.noteLabel.length,
										),
										italics: true,
										size: 20,
										color: ACCENT,
										font: FONT,
									}),
									link(job.noteLabel, job.noteUrl),
								]
							: [
									new TextRun({
										text: job.note,
										italics: true,
										size: 20,
										color: ACCENT,
										font: FONT,
									}),
								],
					}),
				);
			}

			for (const highlight of job.highlights) {
				children.push(
					bullet([bold(highlight.label + ": "), norm(highlight.text)]),
				);
			}
		}
	}

	// Projects
	if (projects.length) {
		children.push(sectionHeading("Personal Projects"));
		for (const project of projects) {
			children.push(
				titledLinkParagraph(project.title, project.label, project.url),
				bullet(project.description),
			);
		}
	}

	// Skills
	if (skills.length) {
		children.push(sectionHeading("Technical Skills"));
		for (const group of skills) {
			children.push(
				bullet([bold(group.group + ": "), norm(group.items.join(", "))]),
			);
		}
	}

	// Education
	if (education.length) {
		children.push(sectionHeading("Education"));
		for (const entry of education) {
			children.push(
				new Paragraph({
					spacing: { after: 10 },
					children: [
						new TextRun({
							text: entry.degree,
							bold: true,
							size: 20,
							color: NAVY,
							font: FONT,
						}),
					],
				}),
				new Paragraph({
					spacing: { after: 0 },
					children: [
						new TextRun({
							text: entry.note,
							size: 20,
							color: ACCENT,
							font: FONT,
							italics: true,
						}),
					],
				}),
			);
		}
	}

	return children;
};

const doc = new Document({
	numbering: {
		config: [
			{
				reference: "bullet-list",
				levels: [
					{
						level: 0,
						format: LevelFormat.BULLET,
						text: "\u2022",
						alignment: AlignmentType.LEFT,
						style: {
							paragraph: {
								indent: {
									left: convertInchesToTwip(0.22),
									hanging: convertInchesToTwip(0.16),
								},
							},
						},
					},
				],
			},
		],
	},
	sections: [
		{
			properties: {
				page: {
					size: { width: 12240, height: 15840 }, // US Letter
					margin: {
						top: convertInchesToTwip(0.18),
						bottom: convertInchesToTwip(0.15),
						left: convertInchesToTwip(0.7),
						right: convertInchesToTwip(0.7),
					},
				},
			},
			children: buildChildren(),
		},
	],
});

const buildExportPayload = () => ({
	...content,
	generatedAt: new Date().toISOString(),
});

const outputDir = __dirname;
const docxPath = path.join(outputDir, `${FILE_STEM}.docx`);
const pdfPath = path.join(outputDir, `${FILE_STEM}.pdf`);
const jsonPath = path.join(outputDir, "resume.json");
const imagePrefix = path.join(outputDir, FILE_STEM);

Packer.toBuffer(doc).then(
	(buffer) => {
		try {
			fs.writeFileSync(docxPath, buffer);
			console.log("DOCX generated.");
		} catch (err) {
			console.error("Failed to write DOCX file.");
			console.error(err.message);
			return; // no docs no need to continue
		}

		try {
			execSync(
				`libreoffice --headless --convert-to pdf "${docxPath}" --outdir "${outputDir}"`,
				{
					stdio: "inherit",
				},
			);
			console.log("PDF generated.");
		} catch (err) {
			console.error("Failed to convert DOCX to PDF.");
			console.error(err.message);
			return; // no pdf no image
		}

		try {
			//? each page will get an image for it
			// execSync(`pdftoppm -png -r 150 "${pdfPath}" "${imagePrefix}"`, {
			// 	stdio: "inherit",
			// });
			//? here only one image file
			//? i will make sure it only will has one page so one file is fine
			execSync(
				`pdftoppm -png -r 150 -singlefile "${pdfPath}" "${imagePrefix}"`,
				{
					stdio: "inherit",
				},
			);
			console.log("Preview image generated.");
		} catch (err) {
			console.error("Failed to convert PDF to image.");
			console.error(err.message);
		}
		// create the json
		try {
			fs.writeFileSync(
				jsonPath,
				JSON.stringify(buildExportPayload(), null, 2) + "\n",
			);
			console.log("JSON export generated.");
		} catch (err) {
			console.error("Failed to write JSON export.");
			console.error(err.message);
		}
	},
	//
);
