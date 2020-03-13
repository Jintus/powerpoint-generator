const FileSystem = require('fs');
const path = require('path');
const PPTX = require('nodejs-pptx');
const Excel = require('exceljs');

// Uncomment when running the app in dev mode (& comment next one)
// const assetsFolder = path.join(process.cwd(), 'assets');
// Uncomment when packaging the app (& comment previous line)
const assetsFolder = path.join(path.dirname(process.execPath), 'assets');
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

async function editPowerPoint(kpis, filename, outputFilename) {
    console.log(`Generating PowerPoint from ${filename}`);
    const pptx = new PPTX.Composer();
    await pptx.load(filename);
    await pptx.compose(async pres => {
        Object.values(pres.powerPointFactory.slides).forEach(slide => {
            let stringifiedSlideContent = JSON.stringify(slide.content['p:sld']);
            Object.entries(kpis).forEach(([kpi, value]) => {
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
        });
    });
    await pptx.save(outputFilename);
    console.log(`PowerPoint generated and saved at ${outputFilename}`);
}
