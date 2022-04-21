/**
 * Global vars
 */
let ss = SpreadsheetApp.getActiveSpreadsheet()
    , ui = SpreadsheetApp.getUi();

/**
 * Adds a customer to the sheet and reports ASE migration status
 */
function addCustomer() {

    // add customer dialog
    let dialog = ui.prompt(
        'Add Customer'
        , 'Account Switch Key ("accountId:contractTypeId")'
        , ui.ButtonSet.OK_CANCEL);

    if (dialog.getSelectedButton() == ui.Button.OK && dialog.getResponseText()) {
        let ask = dialog.getResponseText()
        let askSlim = ask.split(':');
        let customerName = identifyAccount(askSlim[0]);

        // insert the new customer sheet
        ss.insertSheet(customerName);
        SpreadsheetApp.setActiveSheet(ss.getSheetByName(customerName));

        // build full object and assign columns
        let policyObj = buildFullObj(ask);

        // define headers and their widths
        let columnHeaders = [
            'Configuration'
            , 'Policy ID'
            , 'Policy Name'
            , 'Current State'
            , 'Eval'
            , 'Evaluation Type'
            , 'Expires'
            , 'Mode'
        ];
        let columnWidths = [
            180
            , 130
            , 270
            , 210
            , 100
            , 280
            , 170
            , 100
        ]

        // write header columns and set basic styling
        writeColumnHeaders(columnHeaders, columnWidths, customerName);

        // write policy data
        columnHeaders.forEach((header, index) => writePolicyData(index, header, policyObj));
    }
}

/**
 * Writes the array values to header cells of each column
 */
function writeColumnHeaders(headers, widths, account) {
    let customerSheet = ss.getSheetByName(account);

    // write our header values and style
    widths.forEach((item, index) => customerSheet.setColumnWidth(index + 1, item));
    customerSheet.getRange(1, 1, 1, headers.length)
        .setBackground('#d0e0e3')
        .setValues([headers]);
}

/**
 * Writes the data for each config by column to maintain order
 */
function writePolicyData(column, header, policies) {
    let active = ss.getActiveSheet();
    let lastRow = policies.length + 1;

    switch (header) {
        case 'Configuration':
            policies.forEach((policy, index) => {
                writeCell(column, index, policy.configName)
            });
            break;
        case 'Policy ID':
            policies.forEach((policy, index) => {
                writeCell(column, index, policy.policyId)
            });
            break;
        case 'Policy Name':
            policies.forEach((policy, index) => {
                writeCell(column, index, policy.policyName)
            });
            break;
        case 'Current State':
            policies.forEach((policy, index) => {
                writeCell(column, index, policy.current)
            });
            break;
        case 'Eval':
            active.getRange(1, column + 1, lastRow).setHorizontalAlignment('center');
            policies.forEach((policy, index) => {
                writeCell(column, index, policy.eval)
            });
            break;
        case 'Evaluation Type':
            policies.forEach((policy, index) => {
                writeCell(column, index, policy.evaluating)
            });
            break;
        case 'Expires':
            active.getRange(1, column + 1, lastRow).setHorizontalAlignment('center');
            policies.forEach((policy, index) => {
                writeCell(column, index, policy.expires)
            });
            break;
        case 'Mode':
            active.getRange(1, column + 1, lastRow).setHorizontalAlignment('center');
            policies.forEach((policy, index) => {
                writeCell(column, index, policy.mode)
            });
            break;
    }
}

/**
 * Writes the value for each policy in respective row per column
 */
function writeCell(column, row, value) {
    let active = ss.getActiveSheet();

    // write the value (shift values to account for header and 0 based array)
    active.getRange(row + 2, column + 1)
        .setValue(value);
}

/**
 * Builds the full object of security configs and their state
 */
function buildFullObj(accountSwitchKey) {
    let configurations = getSecConfigs(accountSwitchKey);
    let configsComplete = buildSecPolicies(configurations, accountSwitchKey);

    // return fully built array of policy data
    return configsComplete;
}


/**
 * Builds the custom data object for each policy and it's status
 */
function buildSecPolicies(configs, ask) {
    let secPolicies = [];

    configs.forEach(config => {
        let policies = fetchSecPolicies(config, ask);

        policies.forEach(policy => {
            let mode = fetchSecPolicyMode({ ...policy, ...config }, ask)
            let obj = {
                ...config
                , ...policy
                , ...mode
            };

            secPolicies.push(obj);
        });
    });

    return secPolicies;
}

/**
 * Hanldes the actual fetch for each policy
 */
function fetchSecPolicies(config, ask) {
    const eggas = EdgeGridGAS.init({ file: '.edgerc', section: 'default' });

    let res = eggas.auth({
        path: `/appsec/v1/configs/${config.configId}/versions/${config.productionVersion}/security-policies`
        , method: 'GET'
        , headers: {
            'Accept': 'application/json'
        }
        , qs: {
            accountSwitchKey: `${ask}`
        }
    }).send();

    let json = JSON.parse(res);

    return json.policies;

}

/**
 * Hanldes the fetch for each policy mode, and extends into our custom object
 */
function fetchSecPolicyMode(policy, ask) {
    const eggas = EdgeGridGAS.init({ file: '.edgerc', section: 'default' });

    let res = eggas.auth({
        path: `/appsec/v1/configs/${policy.configId}/versions/${policy.productionVersion}/security-policies/${policy.policyId}/mode`
        , method: 'GET'
        , headers: {
            'Accept': 'application/json'
        }
        , qs: {
            accountSwitchKey: `${ask}`
        }
    }).send();

    let fullMode = JSON.parse(res);

    // some mode attributes are missing if not in an active eval
    fullMode.evaluating = fullMode.evaluating || '';
    fullMode.expires = fullMode.expires || '';

    return fullMode;

}

/**
 * Gets all of the account security configs (prod only)
 */
function getSecConfigs(accountSwitchKey) {
    const eggas = EdgeGridGAS.init({ file: '.edgerc', section: 'default' });

    let res = eggas.auth({
        path: '/appsec/v1/configs'
        , method: 'GET'
        , headers: {
            'Accept': 'application/json'
        }
        , qs: {
            accountSwitchKey: `${accountSwitchKey}`
        }
    }).send();

    let json = JSON.parse(res);

    // drop any config that doesn't have a production version
    let onlyProdConfigs = json.configurations.filter(obj => obj.hasOwnProperty("productionVersion"));
    let secConfigs = [];

    // build each config object and push to array
    onlyProdConfigs.forEach(config => {
        let secConfig = {
            configId: config.id
            , configName: config.name
            , productionVersion: config.productionVersion
        };
        secConfigs.push(secConfig);
    });

    // built array of prod only configs and needed attributes
    return secConfigs;

}

/**
 * Takes accountSwitchKey, and converts to account name
 */
function identifyAccount(accountSwitchKey) {
    const eggas = EdgeGridGAS.init({ file: '.edgerc', section: 'default' });
    let res = eggas.auth({
        path: 'identity-management/v2/api-clients/self/account-switch-keys'
        , method: 'GET'
        , headers: {
            'Accept': 'application/json'
        }
        , qs: {
            search: `${accountSwitchKey}`
        }
    }).send();

    let json = JSON.parse(res);
    let accountName = json[0].accountName.split('_');
    return accountName[0];

}

/*****************************************************************
 * 
 * 
 * *****************     Auth & Admin Methods     ****************   
 * 
 * 
 *****************************************************************/

/**
 * onOpen runs on open of the sheet, or full refresh
 */
function onOpen() {
    // Or DocumentApp or FormApp or DriveApp or SlidesApp
    ui.createMenu('Akamai')
        .addItem('Add Customer', 'addCustomer')
        .addSeparator()
        .addSubMenu(ui.createMenu('Help')
            .addItem('About', 'aboutDialog'))
        .addToUi();
}

function aboutDialog() {
    let help = ui.alert('In the Akamai menu item, select "Add Customer" and enter an Account key. Further instructions can be found at https://github.com/nighthauk/ase-migration. Created by Ryan Hauk. Reach out with support requests.');
}