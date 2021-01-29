const fs = require('fs')
const path = require('path')
const exceljs = require('exceljs')
const cliProgress = require('cli-progress')
const sizeOfImage = require('image-size')
const uniqolor = require('uniqolor')
const fetch = require('node-fetch')
const args = require('minimist')(process.argv.slice(2));

const PROGRESS_BAR = Symbol("PROGRESS_BAR");

const readline = require('readline');

const rl = readline.createInterface({
	input: process.stdin,
	output: process.stdout
});

(async () => {
	if (args['h'] || args['help']) {
		console.log(`
Usage:
    ${process.argv0} -h
    ${process.argv0} --help
    ${process.argv0} [-v] [-o|--output=facebook-messages.xlsx] [-i|--in-place] [-f|--force] [--no-images] [--] [path]
    
    -o, --output file.xlsx
		filename to write workbook to
    
    -i, --in-place
    	load file before operation
    
	-f, --force
		force overwrite file even if exists
		
	--no-images
		Reduce XLSX weight by not embedding images
`)
	} else {
		const verbose = !!args['v'] || !!args['verbose']
		const output = args['output'] || args['o'] || "facebook-messages.xlsx"
		const inPlace = !!args['in-place'] || !!args['i']
		const force = !!args['force'] || !!args['f']
		const noImages = args['images'] === false

		const workbook = new exceljs.Workbook()
		if (inPlace) await workbook.xlsx.readFile(output)

		if (!force && !inPlace) {
			if (fs.existsSync(output)) {
				const answer = await new Promise(resolve => rl.question("File already exists. Continue overwriting? [y/N] ", resolve))
				if (!answer.match(/y(es)?/i)) {
					console.error("File exists, not overwriting")
					process.exit(1)
					return
				}
			}
		}

		if (verbose) {
			console.log("Got arguments:", args);
		}

		const directory = args['_'][0] || "."

		const sheet = workbook.addWorksheet(path.basename(directory))
		// noinspection JSValidateTypes
		sheet.columns = [
			{key: 'id', header: "#", width: 5},
			{key: 'date', header: "Date", width: 14},
			{key: 'type', header: "Message type", hidden: true},
			{key: 'sender', header: "Sender", width: 20},
			{key: 'message', header: "Message", width: 50},
		]

		for (let i = 1; i <= 5; i++) sheet.getColumn(i).alignment = {vertical: 'top'}

		sheet.getColumn('date').numFmt = "m.d. hh:mm:ss"
		sheet.getColumn('message').alignment = {wrapText: true}

		sheet.getRow(1).alignment = {vertical: 'bottom'}
		sheet.getRow(1).font = {bold: true}
		sheet.getRow(1).border = {bottom: {style: 'thick'}}

		let rowIndex = 2
		let lastRow = undefined

		async function addImage(cell, uri, options = {}) {
			if (noImages || uri.match(/^https?:\/\/interncache-prn.fbcdn.net\//)) {
				if (options.hyperlinks) {
					cell.value = {
						text: "❎ Broken thumbnail",
						hyperlink: uri.match(/^(https?|file):\/\//) ? uri : getFileLink(uri),
						...options.hyperlinks,
					}
				}
				return
			}

			const image = uri.match(/^(https?|file):\/\//) ? workbook.addImage({
				buffer: (await fetch(uri)).buffer(),
				extension: path.posix.extname(uri),
			}) : workbook.addImage({
				filename: getFilename(uri),
				extension: path.posix.extname(uri),
			})

			if (!cell.value) {
				try {
					const {width, height} = sizeOfImage(getFilename(uri))
					sheet.getRow(rowIndex - 1).height = 263 / width * height
				} catch (e) {
					sheet.getRow(rowIndex - 1).height = 100
				}
			}

			sheet.addImage(image, {
				tl: {col: sheet.getColumn('message').number - 1, row: rowIndex - 2},
				br: {col: sheet.getColumn('message').number - 0, row: rowIndex - 1},
				editAs: 'oneCell',
				hyperlinks: {
					hyperlink: getFileLink(uri),
					tooltip: "Image 🖼",
				},
				...options,
			})
		}

		const jsons = []

		const multibar = new cliProgress.MultiBar()

		const totalProgressBar = multibar.create(0, 0, "Total")

		for await (const dirEntry of await fs.promises.opendir(directory)) {
			if (!dirEntry.isFile) continue

			const filename = dirEntry.name
			const filePath = path.join(directory, filename)

			const filenameMatch = /^message_(\d+)\.json$/.exec(filename)
			if (!filenameMatch) continue

			const n = parseInt(filenameMatch[1])

			// if (jsons[n]) throw new Error(`${filePath} seems to be parsed already`)

			const text = await fs.promises.readFile(filePath)

			jsons[n] = JSON.parse(text)
			jsons[n][PROGRESS_BAR] = multibar.create(jsons[n].messages.length, 0, filename)
		}

		totalProgressBar.start(jsons.length, 0, "Total")

		for (let j = jsons.length - 1; j >= 0; j--) {
			totalProgressBar.update(jsons.length - j)

			const json = jsons[j]

			if (!json) continue

			fixObject(json, {
				participants: false,
				messages: false,
				title: true,
				is_still_participant: false,
				thread_type: undefined,
				thread_path: undefined,
			})
			const {participants, messages, title, thread_path} = json

			function getFileLink(uri) {
				return `file:///${path.resolve(process.cwd(), getFilename(uri)).replaceAll(path.delimiter, path.posix.delimiter).replace(/^\//, '')}`
			}

			function getFilename(uri) {
				return path.join(directory, '..', '..', '..', uri.replaceAll(path.posix.delimiter, path.delimiter))
			}

			sheet.name = title

			for (const participant of participants) {
				fixObject(participant, {
					name: true,
				})
				const {name} = participant

				if (!sheet.columns.some(({key}) => key === `reaction/${name}`)) {
					sheet.columns = [
						...sheet.columns,
						{key: `reaction/${name}`, header: name, width: 4},
					]
					sheet.getColumn(`reaction/${name}`).alignment = {horizontal: 'center'}
					setColor(sheet.getRow(1).getCell(`reaction/${name}`), name)
					sheet.getRow(1).getCell(`reaction/${name}`).alignment = {textRotation: 90}
				}
			}

			const progressBar = json[PROGRESS_BAR]
			progressBar.start(messages.length)

			for (let i = messages.length - 1; i >= 0; i--) {
				progressBar.update(messages.length - i)
				const message = messages[i]
				fixObject(message, {
					sender_name: true,
					timestamp_ms: undefined,
					content: true,
					photos: false,
					reactions: false,
					share: false,
					files: false,
					type: undefined,
					sticker: false,
					gifs: false,
					videos: false,
					call_duration: undefined,
					audio_files: false,
					users: false,
				})

				function newCell() {
					const row = lastRow = sheet.getRow(rowIndex++)
					row.getCell('id').value = Number(i)
					row.getCell('date').value = new Date(message['timestamp_ms'])
					row.getCell('type').value = String(message['type'])
					row.getCell('sender').value = String(message['sender_name'])
					setColor(row.getCell('sender'), message['sender_name'])

					return row.getCell('message')
				}

				if (message.content || message.share) {
					const cell = newCell()

					if (message.share) {
						fixObject(message.share, {
							link: undefined,
							share_text: true,
						})

						const {link, share_text} = message.share

						cell.value = {
							text: share_text || message.content || link,
							hyperlink: link,
						}
					} else {
						cell.value = message.content
					}

					if (message.call_duration !== undefined) {
						cell.value += ` (${message.call_duration}s)`
					}
				}

				if (message.photos) {
					for (const photo of message.photos) {
						fixObject(photo, {
							uri: undefined,
							creation_timestamp: undefined,
						})

						const {uri} = photo

						await addImage(newCell(), uri)
					}
				}

				if (message.videos) {
					for (const video of message.videos) {
						fixObject(video, {
							uri: undefined,
							creation_timestamp: undefined,
							thumbnail: false,
						})

						const {uri} = video

						if (video.thumbnail) {
							fixObject(video.thumbnail, {
								uri: undefined,
							})

							const {uri: thumbUri} = video.thumbnail

							await addImage(newCell(), thumbUri, {
								hyperlinks: {
									hyperlink: getFileLink(uri),
									tooltip: "Video ▶",
								},
							})
						} else {
							newCell().value = {
								text: "Video ▶",
								hyperlink: getFileLink(uri),
							}
						}
					}
				}

				if (message.sticker) {
					fixObject(message.sticker, {
						uri: undefined,
					})

					const {uri} = message.sticker

					await addImage(newCell(), uri)
				}

				if (message.gifs) {
					for (const gif of message.gifs) {
						fixObject(gif, {
							uri: undefined,
						})

						const {uri} = gif

						await addImage(newCell(), uri)
					}
				}

				if (message.audio_files) {
					for (const file of message.audio_files) {
						fixObject(file, {
							uri: undefined,
							creation_timestamp: undefined,
						})
						const {uri} = file

						newCell().value = {
							text: "▶ 🎧 Audio clip",
							hyperlink: getFileLink(uri),
						}
					}
				}

				if (message.files) {
					for (const file of message.files) {
						fixObject(file, {
							uri: undefined,
							creation_timestamp: undefined,
						})
						const {uri} = file

						newCell().value = {
							text: path.posix.basename(uri),
							hyperlink: getFileLink(uri),
						}
					}
				}

				if (lastRow) {
					if (message.reactions) {
						for (const o of message.reactions) {
							fixObject(o, {
								reaction: true,
								actor: true,
							})
							const {reaction, actor} = o
							lastRow.getCell(`reaction/${actor}`).value = (lastRow.getCell(`reaction/${actor}`).value || "") + reaction
							setColor(sheet.lastRow.getCell(`reaction/${actor}`), actor)
						}
					}

					lastRow.border = {bottom: {style: 'thin'}}
				} else {
					console.warn("The following message caused no rows: %o", message)
				}
			}

			if (lastRow) {
				lastRow.addPageBreak()
				lastRow.border = {bottom: {style: 'double'}}
			}

			// progressBar.stop()
		}

		multibar.stop()

		console.log("Writing to %s...", output)
		await workbook.xlsx.writeFile(output)
	}
})().then(() => {
	console.log("Done")
	process.exit(0)
}, e => {
	console.error("Main thread failed: %o", e)
	process.exit(1)
})

function fixObject(o, properties) {
	for (let p in properties) {
		// noinspection JSUnfilteredForInLoop
		if (properties[p] === true) {
			// noinspection JSUnfilteredForInLoop
			if (o[p]) o[p] = fixEncoding(o[p])
		}
	}

	for (let p of Object.getOwnPropertyNames(o)) {
		if (!properties.hasOwnProperty(p)) {
			console.warn("The following object has an unexpected property %O: %o", p, o)
		}
	}
}

function fixEncoding(string, properties) {
	if (properties) throw new Error("Deprecated")

	if (properties) {
		for (let property of properties) {
			if (string[property]) string[property] = fixEncoding(string[property])
		}
	} else {
		return Buffer.from(string, 'latin1').toString('utf8')
	}
}

function setColor(cell, seed) {
	const {color} = uniqolor(seed + (args.salt || ""), {lightness: 80})
	cell.fill = {
		type: 'pattern',
		pattern: 'solid',
		fgColor: {argb: "FF" + color.slice(1).toUpperCase()},
	}
}
