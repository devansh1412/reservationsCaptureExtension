let button1 = document.getElementById('fetchJSON');
let button2 = document.getElementById('datatoExcel');
let button3 = document.getElementById('datatoAPI');
let textArea1 = document.getElementById('accountsDone');
let textArea2 = document.getElementById('accountsLeft');
let jsonData = "";
let member = "";
let accountDetails = {"00016014043":"70k as", "00201778681":"33kJan", "00201726622":"6k Aug", "00201739619":"11k", "00053001682":"7k sspp", "00201866254":"20k May", 
  "00043001472":"20k april", "00203234906":"10k Oct", "00004004009":"12k aug", "00034001704":"10k feb", "00032002734":"20k nov", "00201760020":"7k nov ss", 
  "00005001488":"40k dheer","00201470278":"20k sid","00201734319":"35k sid","00012012699":"20k pp jan","00201629989":"105k", "00022011229":"130k april", "00009011371":"74k",
  "00034002491":"65k rahul", "00201695994":"10k sid","00201021583":"63k dec","00013005961":"17k ssgs","00011006842":"117k ppmp","00009900015":"12k manish","00014015267":"81k",
  "00003004718":"80k sid","00012005521":"36k abhi","00005012234":"63k may","00006011737":"46k","00201251583":"65k joseph","00009900037":"35k suresh","00038007067":"yamini14k",
  "00201602689":"45k","00030001367":"20k Madhu","00050003048":"25k piyush","00014011105":"16k yamini","00201687454":"20k moloy","00203958164":"136k"}

button1.addEventListener('click', async () => {
  let [tab] = await chrome.tabs.query({active : true, currentWindow : true});
  chrome.scripting.executeScript({
    target : {tabId: tab.id},
    function : fetchJson,
  }); 
});

chrome.runtime.onMessage.addListener((request, sender, sendResponse) => {  
  if(request.reservationData === undefined){
    textArea1.textContent = "Please visit Worldmark site and login to an account. This extension wont work otherwise.";  
  }
  else{
    jsonData = JSON.parse(request.reservationData);
    member = JSON.parse(request.memberData);
    let id = member.memberNumber;
    let creds = member.pointsAvailable;
    console.log(creds);
    chrome.storage.local.get('accounts', (result) => {
      const accountlist = result.accounts || {};
      accountlist[id] = jsonData;
      chrome.storage.local.set({ accounts: accountlist }, () => {
        console.log('Account added:',id);    
        chrome.storage.local.get('accounts', (result) => {
          setContent(result.accounts);
        });
      });
    });
    chrome.storage.local.get('credits', (result) => {
      const creditList = result.credits || {};
      creditList[id] = creds;
      chrome.storage.local.set({ credits : creditList }, () => {
        console.log('Credits added:',id);
      });
    });
  }
});

function setContent(accounts){
  let accountsDone = [];
  let accountsLeft = [];
  for(let acc in accountDetails){
    if (accounts[acc] !== undefined) {
      accountsDone.push(accountDetails[acc]);
    } 
    else {
      accountsLeft.push(accountDetails[acc]);
    }
  }
  let output1= "Accounts stored:-\n\n"+accountsDone.join(', ');
  let output2 = "Accounts left:- \n\n"+accountsLeft.join(', ');
  textArea1.textContent = output1;
  textArea2.textContent = output2;
}

function fetchJson(){
  let reservationData = sessionStorage.allReservations;
  let memberData = sessionStorage.contentViewAnalytics;
  chrome.runtime.sendMessage({reservationData, memberData});
}

button2.addEventListener('click', async () => {
  chrome.storage.local.get('accounts', (result) => {
    if (result.accounts) {
      let excelData = [];
      for(let i in result.accounts){
          let reservations = result.accounts[i].reservationSummary;
          for (let j in reservations){
            reservations[j].accountName=accountDetails[i];
            excelData.push(reservations[j]);
          }
      }
      let creditsData = [];
      chrome.storage.local.get('credits', (result) => {
        creditsData.push(["Account", "creds"]);
        for(let i in result.credits){
          creditsData.push([ i, result.credits[i] ]);
        }
      });
      let data = formatData(excelData);
      console.log(creditsData);
      saveToExcel(creditsData,data);
      //chrome.storage.local.clear();
    } 
    else {
      textArea1.textContent = 'No data found.';
    }
  });
});

button3.addEventListener('click', async () => {
  chrome.storage.local.get('accounts', (result) => {
    if (result.accounts) {
      let dataToSend = result.accounts;
      fetch('http://localhost:8080/api/endpoint', {
          method: 'POST',
          headers: {
              'Content-Type': 'application/json'
          },
          body: JSON.stringify(dataToSend)
      })
      .then(response => {
        console.log(response);
      })
      .catch((error) => {
          console.error('Error:', error);
      });
    } 
    else {
      textArea1.textContent = 'No data found.';
    }
  });
});

function formatData(data){
  for (let i in data) {
    data[i].bookedBy = data[i].bookedBy.firstName+" "+data[i].bookedBy.lastName;
    data[i].resort = data[i].resort.resortName;
    data[i].unit = data[i].unit.name;
    data[i].overlapping = data[i].overlapping.eligible;
    data[i].traveler = data[i].traveler.firstName+" "+data[i].traveler.lastName;
    data[i].costDollars = data[i].costDollars.amount;
    data[i].costPoints = data[i].costPoints.amount;
    data[i].costCurrencies = data[i].costCurrencies.amount;
  }
  return data;
}

function saveToExcel(creditsData,data){
  console.log(creditsData);
  const worksheet1 = XLSX.utils.aoa_to_sheet(creditsData);
  console.log(worksheet1);
  const newOrder = ['reservationId', 'confirmationId', 'accountName', 'resort', 'checkInDate', 'checkOutDate', 'cancellationId', 'unit', 'costPoints', 'costDollars', 'traveler', 'bookedBy', 'bookedDate','bookingChannel','lengthOfStay','overlapping']; 
  const newWS = XLSX.utils.json_to_sheet(data, {header:newOrder});
  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, worksheet1, "Sheet2");
  XLSX.utils.book_append_sheet(workbook, newWS, "Sheet1");
  XLSX.writeFile(workbook, "AllReservations.xlsx");
}