
let bids = JSON.parse(localStorage.getItem('bids')||'[]');

function render(){
document.querySelector('#tbl tbody').innerHTML =
bids.map(b=>`<tr><td>${b.bidNo}</td><td>${b.project}</td><td>${b.status}</td></tr>`).join('');

document.getElementById('stats').innerHTML =
`Total Bids: ${bids.length}`;

localStorage.setItem('bids',JSON.stringify(bids));
}
function addBid(){
const bidNo=document.getElementById('bidNo').value;
const project=document.getElementById('projectName').value;

if(!bidNo||!project)return;

bids.push({
bidNo,
project,
status:'Technical Document Collection'
});
render();
}
render();
