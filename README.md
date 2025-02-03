# Excel-Based TwinSAFE Configuration Generator

This repository contains an Excel spreadsheet with VBA scripting capabilities designed to simplify the configuration of TwinSAFE AND function blocks based on a 3x3 grid of sensors and actors.  The spreadsheet allows users to visually map sensors to actors and automatically generates the corresponding XML configuration for TwinSAFE.

## Overview

The Excel sheet provides a user-friendly interface for defining the relationships between sensors and actors.  Users select the appropriate boxes in the 3x3 grid to indicate which sensors should trigger which actors. The VBA script then processes these selections and generates an XML file containing the necessary TwinSAFE AND function block configurations.

## Features

* **Intuitive 3x3 Grid:** A visual representation of the sensor/actor mapping, making configuration straightforward.
* **Automated XML Generation:**  The VBA script automatically generates the TwinSAFE XML configuration, eliminating manual XML editing and reducing the risk of errors.
* **Flexible Mapping:**  Easily configure complex logical relationships between sensors and actors using the grid.
* **Customizable:** The VBA script can be further customized to adapt to specific project requirements (see "Extending the Script" below).

## How to Use

1. **Download the Excel File:** Download the `TwinSAFE-XML-Excel-v1.xlsm` file from this repository.
2. **Open in Excel:** Open the file in Microsoft Excel (ensure macros are enabled).
3. **Configure the Grid:**  In the Excel sheet, you'll find a 3x3 grid.  Each cell in the grid represents a potential connection between a sensor (row) and an actor (column).  Click on a cell to select it.  A selected cell indicates that the corresponding sensor will trigger the corresponding actor.
4. **Configure the Sensors & Actors:** The Sensors and Actors can be renamed. Sensors can also have a "!" character added for negation. (e.g., `!Sensor1`)
5. **Generate XML:** Click the "Generate XML" button (or similar, as implemented in the VBA script).  The script will process the grid selections and create an XML file (e.g., `TwinSAFE_Configuration.xml`).
6. **Import to TwinSAFE:** Import the generated XML file into your TwinSAFE configuration tool.

## VBA Script Explanation (Brief)

The VBA script performs the following key actions:

1. **Reads Grid Selections:**  It iterates through the 3x3 grid and identifies the selected cells.
2. **Creates XML Structure:**  It programmatically builds the XML structure required for the TwinSAFE AND function blocks.
3. **Populates XML with Data:**  It inserts the sensor and actor information into the XML based on the grid selections.
4. **Saves XML to File:**  It saves the generated XML to a file and formats it to be easily readable.

## Extending the Script

The VBA script can be extended to support more complex scenarios, such as:

* **More than 3x3 Grid:** Modify the script to handle larger grid sizes.
* **Different Logic Functions:** Implement other logic functions besides AND (e.g., OR, XOR).
* **Data Validation:** Add data validation to the grid to prevent invalid selections.
* **Error Handling:** Implement robust error handling to manage unexpected inputs.

## Prerequisites

* Microsoft Excel with VBA support.
* Familiarity with TwinSAFE configuration.

## Contributing

Contributions are welcome!  Please open an issue or submit a pull request.

## License

```
BSD Zero Clause License

Copyright (c) 2025 Beckhoff Automation LLC

Permission to use, copy, modify, and/or distribute this software for any purpose with or without fee is hereby granted.

THE SOFTWARE IS PROVIDED "AS IS" AND THE AUTHOR DISCLAIMS ALL WARRANTIES WITH REGARD TO THIS SOFTWARE INCLUDING ALL IMPLIED WARRANTIES OF MERCHANTABILITY AND FITNESS. IN NO EVENT SHALL THE AUTHOR BE LIABLE FOR ANY SPECIAL, DIRECT, INDIRECT, OR CONSEQUENTIAL DAMAGES OR ANY DAMAGES WHATSOEVER RESULTING FROM LOSS OF USE, DATA OR PROFITS, WHETHER IN AN ACTION OF CONTRACT, NEGLIGENCE OR OTHER TORTIOUS ACTION, ARISING OUT OF OR IN CONNECTION WITH THE USE OR PERFORMANCE OF THIS SOFTWARE.
```

## Disclaimer

All sample code provided by Beckhoff Automation LLC are for illustrative purposes only and are provided “as is” and without any warranties, express or implied. Actual implementations in applications will vary significantly. Beckhoff Automation LLC shall have no liability for, and does not waive any rights in relation to, any code samples that it provides or the use of such code samples for any purpose.

