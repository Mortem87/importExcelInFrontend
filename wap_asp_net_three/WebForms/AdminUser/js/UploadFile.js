function ShipmentControlNumberCheck(password)
{

    if (/^[A-Z][0-9]*$/.test(password)) {
        return true;
    } else {
        return false;
    }
}
function ShipmentTypeCheck(password) {

    if (/^SECTION 321$/.test(password)) {
        return true;
    } else {
        return false;
    }
}
function ShipperNameCheck(password) {

    if (/^[A-Z ]*$/.test(password)) {
        return true;
    } else {
        return false;
    }
}
function ShipperAddressCheck(password) {

    if (/^[a-zA-Z0-9 ]*$/.test(password)) {
        return true;
    } else {
        return false;
    }
}
function ShipperCityCheck(password) {

    if (/^[A-Z]+$/.test(password)) {
        return true;
    } else {
        return false;
    }
}
function ShipperCountryCheck(password) {

    if (/^[A-Z]{2}$/.test(password)) {
        return true;
    } else {
        return false;
    }
}
function ShipperStateCheck(password) {

    if (/^[A-Z]{3}$/.test(password)) {
        return true;
    } else {
        return false;
    }
}
function ShipperPostalCheck(password) {

    if (/^[0-9]{5}$/.test(password)) {
        return true;
    } else {
        return false;
    }
}
function ShipperPortofLadingCheck(password) {

    if (/^[a-zA-Z ]*$/.test(password)) {
        return true;
    } else {
        return false;
    }
}
function ConsigneeNameCheck(password) {

    if (/^[a-zA-Z0-9 ]*$/.test(password)) {
        return true;
    } else {
        return false;
    }
}
function ConsigneeAddressCheck(password) {

    if (/^[0-9]+[ ][A-Z0-9][A-Z0-9 ]*$/.test(password)) {
        return true;
    } else {
        return false;
    }
}
function ConsigneeCityCheck(password) {

    if (/^[A-Z ]*$/.test(password)) {
        return true;
    } else {
        return false;
    }
}
function ConsigneeCountryCheck(password) {

    if (/^[A-Z]{2}$/.test(password)) {
        return true;
    } else {
        return false;
    }
}
function ConsigneeStateCheck(password) {

    if (/^[A-Z]{2}$/.test(password)) {
        return true;
    } else {
        return false;
    }
}
function ConsigneePostalCheck(password) {

    if (/^[0-9]{5}$/.test(password)) {
        return true;
    } else {
        return false;
    }
}
function ProductDescriptionCheck(password) {

    if (/^[a-zA-Z0-9 ]*$/.test(password)) {
        return true;
    } else {
        return false;
    }
}
function ProductQtyCheck(password) {

    if (/^[0-9]*$/.test(password)) {
        return true;
    } else {
        return false;
    }
}
function ProductUOMCheck(password) {

    if (/^[A-Z]*$/.test(password)) {
        return true;
    } else {
        return false;
    }
}
function ProductWeightCheck(password) {

    if (/^[0-9]*$/.test(password)) {
        return true;
    } else {
        return false;
    }
}
function ProductUnitofWeightCheck(password) {

    if (/^[A-Z]*$/.test(password)) {
        return true;
    } else {
        return false;
    }
}
function ProductValueCheck(password) {

    if (/^[0-9]*$/.test(password)) {
        return true;
    } else {
        return false;
    }
}
function CustomerReferenceCheck(password) {

    if (/^[A-Z][0-9]*$/.test(password)) {
        return true;
    } else {
        return false;
    }
}
function USPortArriveCheck(password) {

    if (/^[0-9]{4}$/.test(password)) {
        return true;
    } else {
        return false;
    }
}
function FnPortLoadingCheck(password) {

    if (/^[0-9]{5}$/.test(password)) {
        return true;
    } else {
        return false;
    }
}
function FnPortRecieptCheck(password) {

    if (/^[0-9]{5}$/.test(password)) {
        return true;
    } else {
        return false;
    }
}
function OriginCheck(password) {

    if (/^[A-Z]{2}$/.test(password)) {
        return true;
    } else {
        return false;
    }
}
function Upload() {
    //Reference the FileUpload element.
    var fileUpload = document.getElementById("fileUpload");

    //Validate whether File is valid Excel file.
    var regex = /^([a-zA-Z0-9\s_\\.\-:])+(.xls|.xlsx)$/;
    if (regex.test(fileUpload.value.toLowerCase())) {
        if (typeof (FileReader) != "undefined") {
            var reader = new FileReader();

            //For Browsers other than IE.
            if (reader.readAsBinaryString) {
                reader.onload = function (e) {
                    ProcessExcel(e.target.result);
                };
                reader.readAsBinaryString(fileUpload.files[0]);
            } else {
                //For IE Browser.
                reader.onload = function (e) {
                    var data = "";
                    var bytes = new Uint8Array(e.target.result);
                    for (var i = 0; i < bytes.byteLength; i++) {
                        data += String.fromCharCode(bytes[i]);
                    }
                    ProcessExcel(data);
                };
                reader.readAsArrayBuffer(fileUpload.files[0]);
            }
        } else {
            alert("This browser does not support HTML5.");
        }
    } else {
        alert("Please upload a valid Excel file.");
    }
};
function ProcessExcel(data) {
    //Read the Excel File data.
    var workbook = XLSX.read(data, {
        type: 'binary'
    });

    //Fetch the name of First Sheet.
    var firstSheet = workbook.SheetNames[0];

    //Read all rows from First Sheet into an JSON array.
    var excelRows = XLSX.utils.sheet_to_row_object_array(workbook.Sheets[firstSheet]);

    //Create a HTML Table element.
    var table = document.createElement("table");
    table.border = "1";

    //Add the header row.
    var row = table.insertRow(-1);

    //Add the header cells.
    var headerCell = document.createElement("TH");
    headerCell.innerHTML = "Id";
    row.appendChild(headerCell);

    headerCell = document.createElement("TH");
    headerCell.innerHTML = "ShipmentControlNumber";
    row.appendChild(headerCell);

    headerCell = document.createElement("TH");
    headerCell.innerHTML = "ShipmentType";
    row.appendChild(headerCell);

    headerCell = document.createElement("TH");
    headerCell.innerHTML = "ShipperName";
    row.appendChild(headerCell);

    headerCell = document.createElement("TH");
    headerCell.innerHTML = "ShipperAddress";
    row.appendChild(headerCell);

    headerCell = document.createElement("TH");
    headerCell.innerHTML = "ShipperCity";
    row.appendChild(headerCell);

    headerCell = document.createElement("TH");
    headerCell.innerHTML = "ShipperCountry";
    row.appendChild(headerCell);

    headerCell = document.createElement("TH");
    headerCell.innerHTML = "ShipperState";
    row.appendChild(headerCell);

    headerCell = document.createElement("TH");
    headerCell.innerHTML = "ShipperPostal";
    row.appendChild(headerCell);

    headerCell = document.createElement("TH");
    headerCell.innerHTML = "ShipperPortofLading";
    row.appendChild(headerCell);

    headerCell = document.createElement("TH");
    headerCell.innerHTML = "ConsigneeName";
    row.appendChild(headerCell);

    headerCell = document.createElement("TH");
    headerCell.innerHTML = "ConsigneeAddress";
    row.appendChild(headerCell);

    headerCell = document.createElement("TH");
    headerCell.innerHTML = "ConsigneeCity";
    row.appendChild(headerCell);

    headerCell = document.createElement("TH");
    headerCell.innerHTML = "ConsigneeCountry";
    row.appendChild(headerCell);

    headerCell = document.createElement("TH");
    headerCell.innerHTML = "ConsigneeState";
    row.appendChild(headerCell);

    headerCell = document.createElement("TH");
    headerCell.innerHTML = "ConsigneePostal";
    row.appendChild(headerCell);

    headerCell = document.createElement("TH");
    headerCell.innerHTML = "ProductDescription";
    row.appendChild(headerCell);

    headerCell = document.createElement("TH");
    headerCell.innerHTML = "ProductQty";
    row.appendChild(headerCell);

    headerCell = document.createElement("TH");
    headerCell.innerHTML = "ProductUOM";
    row.appendChild(headerCell);

    headerCell = document.createElement("TH");
    headerCell.innerHTML = "ProductWeight";
    row.appendChild(headerCell);

    headerCell = document.createElement("TH");
    headerCell.innerHTML = "ProductUnitofWeight";
    row.appendChild(headerCell);

    headerCell = document.createElement("TH");
    headerCell.innerHTML = "ProductValue";
    row.appendChild(headerCell);

    headerCell = document.createElement("TH");
    headerCell.innerHTML = "CustomerReference";
    row.appendChild(headerCell);

    headerCell = document.createElement("TH");
    headerCell.innerHTML = "USPortArrive";
    row.appendChild(headerCell);

    headerCell = document.createElement("TH");
    headerCell.innerHTML = "FnPortLoading";
    row.appendChild(headerCell);

    headerCell = document.createElement("TH");
    headerCell.innerHTML = "FnPortReciept";
    row.appendChild(headerCell);

    headerCell = document.createElement("TH");
    headerCell.innerHTML = "Origin";
    row.appendChild(headerCell);

    let objsShipments = [];

    //Add the data rows from Excel file.
    for (var i = 0; i < excelRows.length; i++) {
        //Add the data row.
        var row = table.insertRow(-1);

        //Add the data cells.
        var cell = row.insertCell(-1);
        cell.innerHTML = excelRows[i].Id;

        cell = row.insertCell(-1);
        cell.innerHTML = excelRows[i].ShipmentControlNumber;

        cell = row.insertCell(-1);
        cell.innerHTML = excelRows[i].ShipmentType;

        cell = row.insertCell(-1);
        cell.innerHTML = excelRows[i].ShipperName;

        cell = row.insertCell(-1);
        cell.innerHTML = excelRows[i].ShipperAddress;

        cell = row.insertCell(-1);
        cell.innerHTML = excelRows[i].ShipperCity;

        cell = row.insertCell(-1);
        cell.innerHTML = excelRows[i].ShipperCountry;

        cell = row.insertCell(-1);
        cell.innerHTML = excelRows[i].ShipperState;

        cell = row.insertCell(-1);
        cell.innerHTML = excelRows[i].ShipperPostal;

        cell = row.insertCell(-1);
        cell.innerHTML = excelRows[i].ShipperPortofLading;

        cell = row.insertCell(-1);
        cell.innerHTML = excelRows[i].ConsigneeName;

        cell = row.insertCell(-1);
        cell.innerHTML = excelRows[i].ConsigneeAddress;

        cell = row.insertCell(-1);
        cell.innerHTML = excelRows[i].ConsigneeCity;

        cell = row.insertCell(-1);
        cell.innerHTML = excelRows[i].ConsigneeCountry;

        cell = row.insertCell(-1);
        cell.innerHTML = excelRows[i].ConsigneeState;

        cell = row.insertCell(-1);
        cell.innerHTML = excelRows[i].ConsigneePostal;

        cell = row.insertCell(-1);
        cell.innerHTML = excelRows[i].ProductDescription;

        cell = row.insertCell(-1);
        cell.innerHTML = excelRows[i].ProductQty;

        cell = row.insertCell(-1);
        cell.innerHTML = excelRows[i].ProductUOM;

        cell = row.insertCell(-1);
        cell.innerHTML = excelRows[i].ProductWeight;

        cell = row.insertCell(-1);
        cell.innerHTML = excelRows[i].ProductUnitofWeight;

        cell = row.insertCell(-1);
        cell.innerHTML = excelRows[i].ProductValue;

        cell = row.insertCell(-1);
        cell.innerHTML = excelRows[i].CustomerReference;

        cell = row.insertCell(-1);
        cell.innerHTML = excelRows[i].USPortArrive;

        cell = row.insertCell(-1);
        cell.innerHTML = excelRows[i].FnPortLoading;

        cell = row.insertCell(-1);
        cell.innerHTML = excelRows[i].FnPortReciept;

        cell = row.insertCell(-1);
        cell.innerHTML = excelRows[i].Origin;

        let oShipment = new Shipment(
            excelRows[i].Id
            , excelRows[i].ShipmentControlNumber
            , excelRows[i].ShipmentType
            , excelRows[i].ShipperName
            , excelRows[i].ShipperAddress
            , excelRows[i].ShipperCity
            , excelRows[i].ShipperCountry
            , excelRows[i].ShipperState
            , excelRows[i].ShipperPostal
            , excelRows[i].ShipperPortofLading
            , excelRows[i].ConsigneeName
            , excelRows[i].ConsigneeAddress
            , excelRows[i].ConsigneeCity
            , excelRows[i].ConsigneeCountry
            , excelRows[i].ConsigneeState
            , excelRows[i].ConsigneePostal
            , excelRows[i].ProductDescription
            , excelRows[i].ProductQty
            , excelRows[i].ProductUOM
            , excelRows[i].ProductWeight
            , excelRows[i].ProductUnitofWeight
            , excelRows[i].ProductValue
            , excelRows[i].CustomerReference
            , excelRows[i].USPortArrive
            , excelRows[i].FnPortLoading
            , excelRows[i].FnPortReciept
            , excelRows[i].Origin
        );

        objsShipments.push(oShipment);


        

        let shipment_control_number_check = ShipmentControlNumberCheck(oShipment.ShipmentControlNumber);

        if (!shipment_control_number_check) {
            addNode('ERROR IN LINE ' + oShipment.Id + ' DETECTED. VALUE: ' + oShipment.ShipmentControlNumber + ' - ShipmentControlNumber');
        }

        let shipment_type_check = ShipmentTypeCheck(oShipment.ShipmentType);

        if (!shipment_type_check) {
            addNode('ERROR IN LINE ' + oShipment.Id + ' DETECTED. VALUE: ' + oShipment.ShipmentType + ' - ShipperName');
        }

        let shipper_name_check = ShipperNameCheck(oShipment.ShipperName);

        if (!shipper_name_check) {
            addNode('ERROR IN LINE ' + oShipment.Id + ' DETECTED. VALUE: ' + oShipment.ShipperName + ' - ShipperName');
        }

        let shipper_address_check = ShipperAddressCheck(oShipment.ShipperAddress);

        if (!shipper_address_check) {
            addNode('ERROR IN LINE ' + oShipment.Id + ' DETECTED. VALUE: ' + oShipment.ShipperAddress + ' - ShipperAddress');
        }

        let shipper_city_check = ShipperCityCheck(oShipment.ShipperCity);

        if (!shipper_city_check) {
            addNode('ERROR IN LINE ' + oShipment.Id + ' DETECTED. VALUE: ' + oShipment.ShipperCity + ' - ShipperCity');
        }

        let shipper_coutry_check = ShipperCityCheck(oShipment.ShipperCountry);

        if (!shipper_coutry_check) {
            addNode('ERROR IN LINE ' + oShipment.Id + ' DETECTED. VALUE: ' + oShipment.ShipperCountry + ' - ShipperCountry');
        }

        let shipper_state_check = ShipperCityCheck(oShipment.ShipperState);

        if (!shipper_state_check) {
            addNode('ERROR IN LINE ' + oShipment.Id + ' DETECTED. VALUE: ' + oShipment.ShipperState + ' - ShipperState');
        }

        let shipper_postal_check = ShipperPostalCheck(oShipment.ShipperPostal);

        if (!shipper_postal_check) {
            addNode('ERROR IN LINE ' + oShipment.Id + ' DETECTED. VALUE: ' + oShipment.ShipperPostal + ' - ShipperPostal');
        }

        let shipper_port_of_landing_check = ShipperPortofLadingCheck(oShipment.ShipperPortofLading);

        if (!shipper_port_of_landing_check) {
            addNode('ERROR IN LINE ' + oShipment.Id + ' DETECTED. VALUE: ' + oShipment.ShipperPortofLading + ' - ShipperPortofLading');
        }

        let consignee_name_check = ConsigneeNameCheck(oShipment.ConsigneeName);

        if (!consignee_name_check) {
            addNode('ERROR IN LINE ' + oShipment.Id + ' DETECTED. VALUE: ' + oShipment.ConsigneeName + ' - ConsigneeName');
        }

        let consignee_address_check = ConsigneeAddressCheck(oShipment.ConsigneeAddress);

        if (!consignee_address_check) {
            addNode('ERROR IN LINE ' + oShipment.Id + ' DETECTED. VALUE: ' + oShipment.ConsigneeAddress + ' - ConsigneeAddress');
        }

        let consignee_city_check = ConsigneeCityCheck(oShipment.ConsigneeCity);

        if (!consignee_city_check) {
            addNode('ERROR IN LINE ' + oShipment.Id + ' DETECTED. VALUE: ' + oShipment.ConsigneeCity + ' - ConsigneeCity');
        }

        let consignee_country_check = ConsigneeCountryCheck(oShipment.ConsigneeCountry);

        if (!consignee_country_check) {
            addNode('ERROR IN LINE ' + oShipment.Id + ' DETECTED. VALUE: ' + oShipment.ConsigneeCountry + ' - ConsigneeCountry');
        }

        let consignee_state_check = ConsigneeStateCheck(oShipment.ConsigneeState);

        if (!consignee_state_check) {
            addNode('ERROR IN LINE ' + oShipment.Id + ' DETECTED. VALUE: ' + oShipment.ConsigneeState + ' - ConsigneeState');
        }

        let consignee_postal_check = ConsigneePostalCheck(oShipment.ConsigneePostal);

        if (!consignee_postal_check) {
            addNode('ERROR IN LINE ' + oShipment.Id + ' DETECTED. VALUE: ' + oShipment.ConsigneePostal + ' - ConsigneePostal');
        }

        let product_description_check = ProductDescriptionCheck(oShipment.ProductDescription);

        if (!product_description_check) {
            addNode('ERROR IN LINE ' + oShipment.Id + ' DETECTED. VALUE: ' + oShipment.ProductDescription + ' - ProductDescription');
        }

        let product_qty_check = ProductQtyCheck(oShipment.ProductQty);

        if (!product_qty_check) {
            addNode('ERROR IN LINE ' + oShipment.Id + ' DETECTED. VALUE: ' + oShipment.ProductQty + ' - ProductQty');
        }

        let product_uom_check = ProductUOMCheck(oShipment.ProductUOM);

        if (!product_uom_check) {
            addNode('ERROR IN LINE ' + oShipment.Id + ' DETECTED. VALUE: ' + oShipment.ProductUOM + ' - ProductUOM');
        }

        let product_weight_check = ProductWeightCheck(oShipment.ProductWeight);

        if (!product_weight_check) {
            addNode('ERROR IN LINE ' + oShipment.Id + ' DETECTED. VALUE: ' + oShipment.ProductWeight + ' - ProductWeight');
        }

        let product_unit_of_weight_check = ProductUnitofWeightCheck(oShipment.ProductUnitofWeight);

        if (!product_unit_of_weight_check) {
            addNode('ERROR IN LINE ' + oShipment.Id + ' DETECTED. VALUE: ' + oShipment.ProductUnitofWeight + ' - ProductUnitofWeight');
        }

        let product_value_check = ProductValueCheck(oShipment.ProductValue);

        if (!product_value_check) {
            addNode('ERROR IN LINE ' + oShipment.Id + ' DETECTED. VALUE: ' + oShipment.ProductValue + ' - ProductValue');
        }

        let customer_reference_check = CustomerReferenceCheck(oShipment.CustomerReference);

        if (!customer_reference_check) {
            addNode('ERROR IN LINE ' + oShipment.Id + ' DETECTED. VALUE: ' + oShipment.CustomerReference + ' - CustomerReference');
        }

        let us_port_arrive_check = USPortArriveCheck(oShipment.USPortArrive);

        if (!us_port_arrive_check) {
            addNode('ERROR IN LINE ' + oShipment.Id + ' DETECTED. VALUE: ' + oShipment.USPortArrive + ' - USPortArrive');
        }

        let fn_port_loading_check = FnPortLoadingCheck(oShipment.FnPortLoading);

        if (!fn_port_loading_check) {
            addNode('ERROR IN LINE ' + oShipment.Id + ' DETECTED. VALUE: ' + oShipment.FnPortLoading + ' - FnPortLoading');
        }

        let fn_port_reciept_check = FnPortRecieptCheck(oShipment.FnPortReciept);

        if (!fn_port_reciept_check) {
            addNode('ERROR IN LINE ' + oShipment.Id + ' DETECTED. VALUE: ' + oShipment.FnPortReciept + ' - FnPortReciept');
        }

        let origin_check = OriginCheck(oShipment.Origin);

        if (!origin_check) {
            addNode('ERROR IN LINE ' + oShipment.Id + ' DETECTED. VALUE: ' + oShipment.Origin + ' - Origin');
        }

    }

    addNode(objsShipments.length + " LINES PROCESSED.");

    /*
    console.log(objsShipments);

    let ShipmentsJSON = JSON.stringify(objsShipments);

    let ExtendedShipmentsJSON = '{"Shipments":' + ShipmentsJSON + "}";

    console.log(ExtendedShipmentsJSON);
    */
    var dvExcel = document.getElementById("dvExcel");
    dvExcel.innerHTML = "";
    dvExcel.appendChild(table);
};
function addNode(log) {
    var newP = document.createElement("p");
    var textNode = document.createTextNode(log);
    newP.appendChild(textNode);
    document.getElementById("div_results").appendChild(newP);
}
class Shipment {
    Id;
    ShipmentControlNumber;
    ShipmentType;
    ShipperName;
    ShipperAddress;
    ShipperCity;
    ShipperCountry;
    ShipperState;
    ShipperPostal;
    ShipperPortofLading;
    ConsigneeName;
    ConsigneeAddress;
    ConsigneeCity;
    ConsigneeCountry;
    ConsigneeState;
    ConsigneePostal;
    ProductDescription;
    ProductQty;
    ProductUOM;
    ProductWeight;
    ProductUnitofWeight;
    ProductValue;
    CustomerReference;
    USPortArrive;
    FnPortLoading;
    FnPortReciept;
    Origin;
    
    constructor(
        Id
        , ShipmentControlNumber
        , ShipmentType
        , ShipperName
        , ShipperAddress
        , ShipperCity
        , ShipperCountry
        , ShipperState
        , ShipperPostal
        , ShipperPortofLading
        , ConsigneeName
        , ConsigneeAddress
        , ConsigneeCity
        , ConsigneeCountry
        , ConsigneeState
        , ConsigneePostal
        , ProductDescription
        , ProductQty
        , ProductUOM
        , ProductWeight
        , ProductUnitofWeight
        , ProductValue
        , CustomerReference
        , USPortArrive
        , FnPortLoading
        , FnPortReciept
        , Origin
    ) {
        this.Id = Id;
        this.ShipmentControlNumber = ShipmentControlNumber;
        this.ShipmentType = ShipmentType;
        this.ShipperName = ShipperName;
        this.ShipperAddress = ShipperAddress;
        this.ShipperCity = ShipperCity;
        this.ShipperCountry = ShipperCountry;
        this.ShipperState = ShipperState;
        this.ShipperPostal = ShipperPostal;
        this.ShipperPortofLading = ShipperPortofLading;
        this.ConsigneeName = ConsigneeName;
        this.ConsigneeAddress = ConsigneeAddress;
        this.ConsigneeCity = ConsigneeCity;
        this.ConsigneeCountry = ConsigneeCountry;
        this.ConsigneeState = ConsigneeState;
        this.ConsigneePostal = ConsigneePostal;
        this.ProductDescription = ProductDescription;
        this.ProductQty = ProductQty;
        this.ProductUOM = ProductUOM;
        this.ProductWeight = ProductWeight;
        this.ProductUnitofWeight = ProductUnitofWeight;
        this.ProductValue = ProductValue;
        this.CustomerReference = CustomerReference;
        this.USPortArrive = USPortArrive;
        this.FnPortLoading = FnPortLoading;
        this.FnPortReciept = FnPortReciept;
        this.Origin = Origin;
    }
}
