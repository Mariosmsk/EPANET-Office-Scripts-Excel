function main(workbook: ExcelScript.Workbook) {
  let sheet = workbook.getActiveWorksheet();
  let junction_id = 'J1';

  // JUNCTION ID, ELEVATION
  let data = [junction_id, 888];
  // JUNCTION ID, COURDX, COORDY
  let coordinates = [junction_id, 65, 20];

  // Add a junction
  addNodeJunction(sheet, data, coordinates);

  // PIPE ID, FROM_NODE, TO_NODE, LENGTH, DIAMETER, ROUGHNESS, MINORLOSS, STATUS
  let data_pipe = ['P-1', junction_id, '32', 1200, 5, 120, 0, 'CV'];

  // Add a link
  addLinkPipe(sheet, data_pipe);
  return;
}

function addLinkPipe(sheet: ExcelScript.Worksheet, data_pipe: (string | string | string | number | number | number | number | string)[]): void {
  let all_indices = getRowIndex(sheet);
  let pipe_index = all_indices[1];

  let merge_data = '';
  let pipe = '';
  for (pipe in data_pipe) {
    merge_data = merge_data + data_pipe[pipe] + '                             '
  }

  // add junction in section [JUNCTIONS]
  sheet.getRange(`${pipe_index}:${pipe_index}`).insert(ExcelScript.InsertShiftDirection.down);
  let newadd = sheet.getRange(`A${pipe_index}`);
  newadd.setValue(merge_data);
  return;
}

function addNodeJunction(sheet: ExcelScript.Worksheet, data: (string | number)[],
  coordinates: (string | number | number)[]): void {
  let all_indices = getRowIndex(sheet);
  let junction_index = all_indices[0];
  let coord_index = all_indices[2];

  let merge_jdata = '';
  let jdata = '';
  for (jdata in data) {
    merge_jdata = merge_jdata + data[jdata] + '                             '
  }

  // add junction in section [JUNCTIONS]
  sheet.getRange(`${junction_index}:${junction_index}`).insert(ExcelScript.InsertShiftDirection.down);
  let newadd = sheet.getRange(`A${junction_index}`);
  newadd.setValue(merge_jdata);

  let merge_cdata = '';
  let cdata = '';
  for (cdata in data) {
    merge_cdata = merge_cdata + data[cdata] + '                             '
  }

  // add junction coordinates in [COORDINATES]
  sheet.getRange(`${coord_index}:${coord_index}`).insert(ExcelScript.InsertShiftDirection.down);
  let addcoord = sheet.getRange(`A${coord_index}:C${coord_index}`);
  addcoord.setValues([coordinates]);
  return;
}

function getRowIndex(sheet: ExcelScript.Worksheet): (number | number | number)[] {
  let usedRange = sheet.getUsedRange(true);
  const values = usedRange.getValues();
  let i = 0;
  let row_index = 0;
  let junction_index = 0;
  let pipe_index = 0;
  let coord_index = 0;
  let emptyrow = 0;

  for (let row of values) {
    row_index += 1;
    if (row.toString() == ',,,,') {
      emptyrow = row_index;
    }

    if (row.toString() == '[RESERVOIRS],,,,') {
      i += 1;
      junction_index = row_index - 1;
    }

    if (row.toString() == '[PUMPS],,,,') {
      i += 1;
      pipe_index = row_index - 1;
    }

    if (row.toString() == '[VERTICES],,,,') {
      i += 1;
      coord_index = row_index;
    }

  }
  return [junction_index, pipe_index, coord_index];
}
