/*
Fetch List of winners A
Fetch List of winners B
Fetch List of loosers A
Fetch List of loosers B
While concatenating below, persist data about who is winner and who is looser,
also persist data about group.
listOfAllCompanies = winnersA + winnersB + loosersA + loosersB 


ForEach company in listOfAllCompanies : 
company[i] = allColumnsInCompany[i] + ans{1,2,3,4,5,6 + allOtherFeilds} + group + companyName

print listOfAllCompaniesToExcelSheet

*/

const request = require('request');
const json2xls = require('json2xls');
const fs = require('fs');
const cliProgress = require('cli-progress');

function getCompaniesList(glType, indxGrpval) {
	var options = {
		method: 'GET',
		url: `https://api.bseindia.com/BseIndiaAPI/api/MktRGainerLoserData/w?GLtype=${glType}&IndxGrp=group&IndxGrpval=${indxGrpval}&orderby=all`,
		headers: {
			authority: 'api.bseindia.com',
			'sec-ch-ua':
				'" Not;A Brand";v="99", "Google Chrome";v="91", "Chromium";v="91"',
			accept: 'application/json, text/plain, */*',
			'sec-ch-ua-mobile': '?1',
			'user-agent':
				'Mozilla/5.0 (Linux; Android 6.0; Nexus 5 Build/MRA58N) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.114 Mobile Safari/537.36',
			origin: 'https://www.bseindia.com',
			'sec-fetch-site': 'same-site',
			'sec-fetch-mode': 'cors',
			'sec-fetch-dest': 'empty',
			referer: 'https://www.bseindia.com/',
			'accept-language': 'en-GB,en-US;q=0.9,en;q=0.8',
		},
	};
	return new Promise(function (resolve, reject) {
		request(options, function (error, response) {
			if (error) {
				console.log(error);
				reject();
			} else {
				resolve(JSON.parse(response.body).Table);
			}
		});
	});
}

function getCompanyTopData() {
	var options = {
		method: 'GET',
		url: 'https://api.bseindia.com/BseIndiaAPI/api/getScripHeaderData/w?Debtflag=&scripcode=532890&seriesid=',
		headers: {
			authority: 'api.bseindia.com',
			pragma: 'no-cache',
			'cache-control': 'no-cache',
			'sec-ch-ua':
				'" Not;A Brand";v="99", "Google Chrome";v="91", "Chromium";v="91"',
			accept: 'application/json, text/plain, */*',
			'sec-ch-ua-mobile': '?1',
			'user-agent':
				'Mozilla/5.0 (Linux; Android 6.0; Nexus 5 Build/MRA58N) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.114 Mobile Safari/537.36',
			origin: 'https://www.bseindia.com',
			'sec-fetch-site': 'same-site',
			'sec-fetch-mode': 'cors',
			'sec-fetch-dest': 'empty',
			referer: 'https://www.bseindia.com/',
			'accept-language': 'en-GB,en-US;q=0.9,en;q=0.8',
		},
	};
	return new Promise(function (resolve, reject) {
		request(options, function (error, response) {
			if (error) {
				console.log(error);
				reject(error);
			} else {
				resolve(JSON.parse(response.body));
			}
		});
	});
}

function getCompanyMainData(scripCode) {
	var options = {
		method: 'GET',
		url: `https://api.bseindia.com/BseIndiaAPI/api/HighLow/w?Type=EQ&scripcode=${scripCode}`,
		headers: {
			authority: 'api.bseindia.com',
			'sec-ch-ua':
				'" Not;A Brand";v="99", "Google Chrome";v="91", "Chromium";v="91"',
			accept: 'application/json, text/plain, */*',
			'sec-ch-ua-mobile': '?1',
			'user-agent':
				'Mozilla/5.0 (Linux; Android 6.0; Nexus 5 Build/MRA58N) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.114 Mobile Safari/537.36',
			origin: 'https://www.bseindia.com',
			'sec-fetch-site': 'same-site',
			'sec-fetch-mode': 'cors',
			'sec-fetch-dest': 'empty',
			referer: 'https://www.bseindia.com/',
			'accept-language': 'en-GB,en-US;q=0.9,en;q=0.8',
			'if-modified-since': 'Fri, 25 Jun 2021 17:12:37 GMT',
		},
	};
	return new Promise(function (resolve, reject) {
		request(options, function (error, response) {
			if (error) {
				console.log(error);
				reject();
			} else {
				resolve(JSON.parse(response.body));
			}
		});
	});
}

function convertJsonToXls(jsonObject, destinationFileName) {
	return new Promise(function (resolve, reject) {
		var xls = json2xls(jsonObject);
		fs.writeFileSync(destinationFileName + '.xls', xls, 'binary');
		resolve();
	});
}

async function main() {
	glTypesList = ['loser'];
	indxGrpvalList = ['A'];

	allSelectedCompanes = [];

	for (const glType of glTypesList) {
		for (const indxGrpval of indxGrpvalList) {
			console.log(
				'\nFetching gainerLooserType : ' +
					glType +
					' and indexGroup : ' +
					indxGrpval
			);
			companyList = await getCompaniesList(glType, indxGrpval);
			companyList = companyList.slice(1, 4);
			console.log(
				'\nFetched ' +
					companyList.length +
					' companies of' +
					' gainerLooserType : ' +
					glType +
					' and indexGroup : ' +
					indxGrpval
			);
			console.log('Procesing above mentioned categories companies');
			// create a new progress bar instance and use shades_classic theme
			const bar1 = new cliProgress.SingleBar(
				{},
				cliProgress.Presets.shades_classic
			);
			bar1.start(companyList.length, 0);
			for (const company of companyList) {
				companyMainData = await getCompanyMainData(company.scrip_cd);
				requiredCompanyData = {};
				requiredCompanyData.scrip_cd = company.scrip_cd;
				requiredCompanyData.scrip_grp = company['scrip_grp'];
				requiredCompanyData.longName = company['LONG_NAME'];
				requiredCompanyData.ltradert = company['ltradert'];
				requiredCompanyData.change_val = company['change_val'];
				requiredCompanyData.change_percentage =
					company['change_percent'];
				requiredCompanyData.fifty2WkHigh =
					companyMainData.Fifty2WkHigh_adj;
				requiredCompanyData.fifty2WkLow =
					companyMainData.Fifty2WkLow_adj;
				requiredCompanyData.openrate = company.openrate;
				requiredCompanyData.highrate = company.highrate;
				requiredCompanyData.lowrate = company.lowrate;
				requiredCompanyData.date_time = company['dt_tm'];
				requiredCompanyData.url = company.URL;
				allSelectedCompanes.push(requiredCompanyData);
				bar1.increment();
			}
		}
	}
	console.log(allSelectedCompanes[0]);
	console.log('Attempting to create Excel Sheet');
	await convertJsonToXls(allSelectedCompanes, new Date().toString());
	console.log('Excel sheet creation completed');
}

main();
