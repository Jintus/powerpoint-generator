const FileSystem = require('fs');
const path = require('path');
const PPTX = require('nodejs-pptx');
const Excel = require('exceljs');

// Uncomment when running the app in dev mode (& comment next one)
// const assetsFolder = path.join(process.cwd(), 'assets');
// Uncomment when packaging the app (& comment previous line)
const assetsFolder = path.join(path.dirname(process.execPath), 'assets');
const imagesFolder = path.join(assetsFolder, "images");
const configFile = path.join(assetsFolder, 'generator.config.json');
const GeneratorConfig = JSON.parse(FileSystem.readFileSync(configFile));
const KPIExcelConfig = GeneratorConfig.kpiExcelFileConfig;

const excelFile = path.join(assetsFolder, KPIExcelConfig.filename);
const template = path.join(assetsFolder, GeneratorConfig.templatePptFileName);
const output = path.join(assetsFolder, GeneratorConfig.outputPptFileName);

updateTemplateWithKPI(template, excelFile, output);

async function updateTemplateWithKPI(template, excelFile, output) {
    const kpis = await loadKpiFromExcelFile(excelFile);
    editPowerPoint(kpis, template, output);
}

async function loadKpiFromExcelFile(filename) {
    const workbook = new Excel.Workbook();
    console.log(`Loading KPI from ${filename}`);
    await workbook.xlsx.readFile(filename);
    const ws = workbook.getWorksheet(KPIExcelConfig.worksheetIndexOrName);
    const placeholders = ws.getColumn(KPIExcelConfig.placeholdersColumnIndex).values.splice(KPIExcelConfig.dataRowStartIndex);
    const values = ws.getColumn(KPIExcelConfig.valuesColumnIndex).values.splice(KPIExcelConfig.dataRowStartIndex);
    console.log(`KPI loaded`);
    return Object.fromEntries(placeholders.map((placeholder, index) => {
        const value = values[index];
        return [placeholder, typeof value === "object" ? value.result : value];
    }));
}

function addEvolutionImageToSlide(slide, value, positionX, positionY, sizeX, sizeY) {
    let imageFilename = GeneratorConfig.images.stable;
    if (value < 0)
        imageFilename = GeneratorConfig.images.descending;
    else if (value > 0)
        imageFilename = GeneratorConfig.images.ascending;
    console.log(`Adding image ${imageFilename} to ${slide.name}`);
    slide.addImage(image => {
        image
            .file(path.join(imagesFolder, imageFilename))
            .x(positionX)
            .y(positionY)
            .cx(sizeX)
            .cy(sizeY)
    });
}

async function editPowerPoint(kpis, filename, outputFilename) {
    console.log(`Generating PowerPoint from ${filename}`);
    const pptx = new PPTX.Composer();
    await pptx.load(filename);
    await pptx.compose(async pres => {
        Object.values(pres.powerPointFactory.slides).forEach((slide, index) => {
            let stringifiedSlideContent = JSON.stringify(slide.content['p:sld']);
            Object.entries(kpis).forEach(([kpi, value]) => {
                if (kpi.substr(0, 3) === "img") {
                    return;
                }
                if (kpi.substr(0, 2) === "tv") {
                    kpi = kpi.substr(2);
                    if (Number.isFinite(value)) {
                        value = Math.round(value * 10) / 10;
                        if (value > 0) {
                            value = `+${value}`;
                        }
                    }
                }
                value = value === undefined ? "" : value;
                stringifiedSlideContent = stringifiedSlideContent.replace(new RegExp(kpi, 'g'), value);
            });
            slide.content['p:sld'] = JSON.parse(stringifiedSlideContent);
            if (index === 3) {
                addEvolutionImageToSlide(slide, kpis["img_1"], 150, 210, 30, 30);
                addEvolutionImageToSlide(slide, kpis["img_2"], 800, 210, 30, 30);
                addEvolutionImageToSlide(slide, kpis["img_3"], 150, 330, 30, 30);
                addEvolutionImageToSlide(slide, kpis["img_4"], 800, 330, 30, 30);
            }
        });
    });
    await pptx.save(outputFilename);
    console.log(`PowerPoint generated and saved at ${outputFilename}`);
}
