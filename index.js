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

const NAVY = "1F2937";
const ACCENT = "334155";
const RULE = "9CA3AF";

const FONT = "Calibri";

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
			children: [
				// Name
				new Paragraph({
					alignment: AlignmentType.CENTER,
					spacing: { after: 20 },
					children: [
						new TextRun({
							text: "HAMZA SYRAGE",
							bold: true,
							size: 40,
							color: NAVY,
							font: FONT,
						}),
					],
				}),
				new Paragraph({
					alignment: AlignmentType.CENTER,
					spacing: { after: 40 },
					children: [
						new TextRun({
							text: "Front-End Engineer | React & Next.js Specialist",
							size: 22,
							color: ACCENT,
							font: FONT,
						}),
					],
				}),
				new Paragraph({
					alignment: AlignmentType.CENTER,
					spacing: { after: 40 },
					children: [
						new TextRun({
							text: "Damascus, Syria  |  +963 941 845 197  |  hamzasyrage@gmail.com",
							size: 18,
							font: FONT,
							color: "334155",
						}),
						new TextRun({ text: "", break: 1 }),
						link("github.com/HamzaSyrage", "https://github.com/HamzaSyrage"),
						norm("  |  "),
						link("hamza-syrage.is-a.dev", "https://hamza-syrage.is-a.dev"),
						norm("  |  "),
						link(
							"linkedin.com/in/hamzasyrage",
							"https://linkedin.com/in/hamzasyrage",
						),
					],
				}),
				hr(),

				// Summary
				sectionHeading("Professional Summary"),
				new Paragraph({
					spacing: { after: 10 },
					children: [
						norm(
							"Front-end developer with 2+ years of experience shipping production React and Next.js applications for enterprise clients. Specializes in multi-tenant SaaS architecture, real-time systems (WebRTC/WebSockets), and framework-agnostic embeddable widgets. Has taken features from Figma to production on platforms serving thousands of concurrent users, including a live electronic-voting system for a national professional association and a white-label LMS that reskins itself per client at runtime with zero rebuilds. Strong TypeScript fundamentals, comfortable owning architecture decisions independently, and focused on clean, accessible, well-tested UI.",
						),
					],
				}),

				// Experience
				sectionHeading("Professional Experience"),
				new Paragraph({
					spacing: { before: 40, after: 4 },
					tabStops: [{ type: "right", position: convertInchesToTwip(6.5) }],
					children: [
						bold("Frontend Developer - "),
						link("Lucidly", "https://lucidly.ae"),
						new TextRun({
							text: "\tMay 2025 - Present",
							italics: true,
							size: 20,
							color: ACCENT,
							font: FONT,
						}),
					],
				}),
				new Paragraph({
					spacing: { after: 10 },
					children: [
						new TextRun({
							text: "Remote (UAE) - client work delivered for ",
							italics: true,
							size: 20,
							color: ACCENT,
							font: FONT,
						}),
						link("Axenso", "https://axenso.com"),
					],
				}),
				bullet([
					bold("SIFO - Live meeting & e-voting platform: "),
					norm(
						"built the frontend for a congress meeting and formal-election platform for a national pharmacy association, supporting meetings with over 10,000 concurrent attendees. Verified frontend scalability through automated Puppeteer load tests. Designed a dual real-time transport architecture separating WebRTC media (SFU-based) from WebSocket application state, ensuring chat, hand-raise, permissions, and voting remained reliable even when participants experienced media connection issues.",
					),
				]),
				bullet([
					bold("Cube26 - Multi-tenant SaaS LMS: "),
					norm(
						"built the learner-facing app for a white-label training platform serving multiple client organizations from one codebase. Implemented runtime theming (colors, logos, tag palettes resolved per subdomain from a branding API, no rebuild required) and a unified progress-tracking contract across seven content formats (video, audio, PDF, text, spreadsheets, zip archives, image galleries).",
					),
				]),
				bullet([
					bold("RSS Feed Web Component - Embeddable widget: "),
					norm(
						"built a medical/scientific news aggregator end to end as a framework-agnostic native custom element (<rss-feed>), allowing any client site to embed live, per-client-themed content with a single script tag - no iframe or host build step required. Optimized the widget to a 277 KB minified bundle (93 KB gzipped) by avoiding external runtime dependencies and implementing required functionality with native Web APIs.",
					),
				]),
				// bullet([
				// 	bold("Across all projects: "),
				// 	norm(
				// 		"state managed with Jotai + TanStack Query, Figma designs translated into pixel-perfect responsive components, code quality enforced via ESLint, Prettier, and lint-staged.",
				// 	),
				// ]),

				// Projects
				sectionHeading("Personal Projects"),
				new Paragraph({
					spacing: { before: 40, after: 8 },
					children: [
						bold("Portfolio Site - "),
						link("hamza-syrage.is-a.dev", "https://hamza-syrage.is-a.dev"),
					],
				}),
				bullet(
					"Designed and built a full personal site and technical blog on Next.js 16, React 19, TypeScript, Tailwind CSS v4, and MDX, including a custom MDX component system (interactive diagrams, code file trees, callouts) and a library of 30+ hand-built micro-interaction demos using Motion.",
				),
				new Paragraph({
					spacing: { before: 40, after: 8 },
					children: [
						bold("3D Physics-Based Ping Pong Simulation - "),
						link(
							"pinging-and-ponging.vercel.app",
							"https://pinging-and-ponging.vercel.app",
						),
					],
				}),
				bullet(
					"Built a real-time table tennis simulator in Three.js and TypeScript from scratch: RK4 numerical integration, Magnus force and drag modeling, custom collision resolution, bot AI, full scoring/fault rules, and a gyroscope-driven mobile controller with live two-way sync to a debug UI.",
				),
				new Paragraph({
					spacing: { before: 40, after: 8 },
					children: [
						bold("StayBay - "),
						link(
							"github.com/HamzaSyrage/staybay-backend",
							"https://github.com/HamzaSyrage/staybay-backend",
						),
					],
				}),
				bullet(
					"Laravel + Sanctum REST API for an Airbnb-style booking platform: escrow-style hold-balance wallet, a scheduled service that auto-transitions bookings by date and payment state, overlap-safe availability checks with date-range merging, and a dynamic query-filter system for search.",
				),

				// Skills
				sectionHeading("Technical Skills"),
				bullet([
					bold("Frontend: "),
					norm("JavaScript (ES6+), TypeScript, React, Next.js, HTML5, CSS3"),
				]),
				bullet([
					bold("Real-Time & Networking: "),
					norm(
						"WebRTC, WebSockets, RESTful APIs, Web Components / Custom Elements",
					),
				]),
				bullet([
					bold("State, Data & Forms: "),
					norm(
						"TanStack Query, Jotai, Redux Toolkit, Zustand, React Hook Form, Zod",
					),
				]),
				bullet([
					bold("UI & Styling: "),
					norm(
						"Tailwind CSS, Radix UI, Shadcn/ui, Chakra UI, Mantine, Motion/Framer Motion, Sass",
					),
				]),
				bullet([
					bold("3D / Graphics: "),
					norm("Three.js, WebGL, Godot, GDScript, C#"),
				]),
				bullet([
					bold("Testing & Monitoring: "),
					norm("Vitest, Playwright, puppeteer, Sentry"),
				]),
				bullet([
					bold("DevOps & CI/CD: "),
					norm("Linux, Docker, Vercel, AWS, Nginx"),
				]),
				bullet([bold("Backend: "), norm("Node.js, Express.js, PHP, Laravel")]),
				bullet([
					bold("Tooling & Practices: "),
					norm(
						"Git, Figma, Postman, Accessibility (A11y), SEO, Responsive Design",
					),
				]),

				// Education
				sectionHeading("Education"),
				new Paragraph({
					spacing: { after: 10 },
					children: [
						new TextRun({
							text: "Bachelor of Science in Information Technology",
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
							text: "Damascus University - Expected Graduation: 2028",
							size: 20,
							color: ACCENT,
							font: FONT,
							italics: true,
						}),
					],
				}),
			],
		},
	],
});

Packer.toBuffer(doc).then((buffer) => {
	const outputDir = __dirname;
	const docxPath = path.join(
		outputDir,
		"Hamza-Syrage-Frontend-Engineer-Resume.docx",
	);
	const pdfPath = path.join(
		outputDir,
		"Hamza-Syrage-Frontend-Engineer-Resume.pdf",
	);
	const imagePrefix = path.join(
		outputDir,
		"Hamza-Syrage-Frontend-Engineer-Resume",
	);

	Packer.toBuffer(doc).then((buffer) => {
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
				{ stdio: "inherit" },
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
	});
});
