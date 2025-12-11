# KML Generator <img src="https://snipboard.io/rlh6gz.jpg" width="10%" height="00%" align="right" valign="middle"/>
Excel & VBA Based KML Generator for Geolocation Visualization on Google Earth and GIS Tools

<div align="center">

![version](https://img.shields.io/badge/version-1.1-blue.svg)
![status](https://img.shields.io/badge/status-stable-006400.svg)
![excel](https://img.shields.io/badge/Excel-.xlsm-green.svg)
![vba](https://img.shields.io/badge/VBA-macros-yellow.svg)
![kml](https://img.shields.io/badge/KML-export-orange.svg)
![license](https://img.shields.io/badge/license-MIT-black.svg)

</div>

<details>
  <summary>[Open/Close] Table of Contents</summary>

- [KML Generator](#kml-generator-)
  - [Description](#-description)
  - [Technologies](#-technologies)
  - [General Settings](#️-general-settings)
  - [Usage](#-usage)
  - [Repository Structure](#-repository-structure)
  - [Architecture](#️-architecture)
  - [Output Example](#-output-example)
  - [Versions](#-versions)

</details>

## 📄 Description
The **KML Generator** automates the creation of `.kml` files based on coordinate tables maintained in an Excel sheet.

The project includes:

- A **ready-to-use Excel template** (`dados_kml.xlsm`)
- All **VBA modules already imported**
- A **standardized table format** ready for modification
- Built-in validation for **colors**, **coordinate formats**, and **structure integrity**

## 💡 Technologies
- Excel (.xlsm)
- VBA
- KML format
- Google Earth-compatible

## ⚙️ General Settings

### Excel Model Structure
The file **`dados_kml.xlsm`** contains:

- All **VBA modules pre-imported** (`funcoes.bas`, `gerar_kml.bas`)
- A fully prepared structure for edition
- Columns preconfigured for KML generation
- Validation of allowed colors for `IconCor`

### Column Rules and Constraints
You *may*:
- Add or remove columns **before the `Drive` column** (yellow KMZ stage).
- Add metadata fields.

You may **not**:
- Add columns **after `IconCor`**
- Change the order of required columns

Required final structure:

| Link Description | Latitude | Longitude | IconText | Drive | IconCor |
|------------------|----------|-----------|----------|-------|---------|

### Format Requirements
- **Latitude/Longitude format (must include degree symbol °):**
  ```
  37.235000°
  -115.811111°
  ```
- **Drive must be a URL link to a folder**
- **IconCor only accepts implemented script colors**

### IconCor Allowed Colors
```
vermelho
amarelo
verde
branco
```

## 🚀 Usage

### Input Example
| Link Description | Latitude       | Longitude       | IconText | Drive                | IconCor |
|------------------|----------------|-----------------|----------|----------------------|---------|
| Folder           | -22.958273°    | -43.065060°     | Casa 57  | https://drive/folder | azul    |

### Generate KML
```
Developer → Macros → gerar_kml
```

## 📁 Repository Structure
```
/kml-generator
├── dados_kml.xlsm
├── funcoes.bas
├── gerar_kml.bas
├── LICENSE
└── README.md
```

## ⚙️ Architecture
- Input Layer (Excel)
- Processing Layer (VBA)
- Output Layer (KML)

## 🔎 Output Example
```
<Placemark>
  <name>Galpão</name>
  <Point>
    <coordinates>37.235000°, -115.811111°</coordinates>
  </Point>
</Placemark>
```

## 🚧 Versions
### Version 1.1 – 2025/02/11
- Documentation updates
- Added IconCor constraints
- Added Drive rules
- Updated Excel model instructions
