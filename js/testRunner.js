fetch('Calc.test.asp')
  .then(response => response.text())
  .then(str => new window.DOMParser().parseFromString(str, "text/xml"))
  .then(xmldata => {
      visualize(document.getElementById("content"), xmldata);
  });
function visualize(targetElement, xmldata){
    const xmlroot = xmldata.getElementsByTagName("testsuites")[0];
    const testsuites = xmldata.getElementsByTagName("testsuite");
    //h2
    let h2 = document.createElement("h2");
    h2.innerText = xmlroot.getAttribute("name");
    targetElement.appendChild(h2);
    //summary
    //get summary data
    let testsCount = xmlroot.hasAttribute("tests")? xmlroot.getAttribute("tests"): 0;
    let failureCount = xmlroot.hasAttribute("failures")? xmlroot.getAttribute("failures"): 0;
    let errorCount = xmlroot.hasAttribute("errors")? xmlroot.getAttribute("errors"): 0;
    let skippedCount = xmlroot.hasAttribute("skipped")? xmlroot.getAttribute("skipped"): 0;
    let successCount = testsCount - failureCount - errorCount - skippedCount;
    let time = xmlroot.getAttribute("time");
    //append summary
    let summaryDl = document.createElement("dl");
    summaryDl.className = "row";
    targetElement.appendChild(summaryDl);
    //time
    appendDlChild(summaryDl, '<span class="text-body">time</span>', time + '<span class="text-muted">&nbsp;sec</span>');
    //success
    appendDlChild(summaryDl, '<span class="text-success">success</span>', successCount);
    //failure
    appendDlChild(summaryDl, '<span class="text-danger">failure</span>', failureCount);
    //error
    appendDlChild(summaryDl, '<span class="text-warning">error</span>', errorCount);
    //skipped
    appendDlChild(summaryDl, '<span class="text-info">skipped</span>', skippedCount);

    
    for (let index = 0; index < testsuites.length; index++) {
        let testsuite = testsuites[index];
        let h3 = document.createElement("h3");
        h3.innerHTML = testsuite.getAttribute("name");
        targetElement.appendChild(h3);
        
        //summary
        //get summary data
        let testsCount = testsuite.hasAttribute("tests")? testsuite.getAttribute("tests"): 0;
        let failureCount = testsuite.hasAttribute("failures")? testsuite.getAttribute("failures"): 0;
        let errorCount = testsuite.hasAttribute("errors")? testsuite.getAttribute("errors"): 0;
        let skippedCount = testsuite.hasAttribute("skipped")? testsuite.getAttribute("skipped"): 0;
        let successCount = testsCount - failureCount - errorCount - skippedCount;
        let time = testsuite.getAttribute("time");
        //append summary
        let summaryDl = document.createElement("dl");
        summaryDl.className = "row";
        //time
        appendDlChild(summaryDl, '<span class="text-body">time</span>', time + '<span class="text-muted">&nbsp;sec</span>');
        //success
        appendDlChild(summaryDl, '<span class="text-success">success</span>', successCount);
        //failure
        appendDlChild(summaryDl, '<span class="text-danger">failure</span>', failureCount);
        //error
        appendDlChild(summaryDl, '<span class="text-warning">error</span>', errorCount);
        //skipped
        appendDlChild(summaryDl, '<span class="text-info">skipped</span>', skippedCount);

        targetElement.appendChild(summaryDl);

        let testcases = testsuite.getElementsByTagName("testcase");
        let row = document.createElement("dl")
        row.className="row";
        for (let index = 0; index < testcases.length; index++) {
            let testcase = testcases[index];
            let col = document.createElement("dt");
            col.className = "col-md-1";
            col.innerHTML = testcase.hasChildNodes()? '<span class="text-danger">NG</span>':'<span class="text-success">OK</span>';
            row.appendChild(col);
            col = document.createElement("dd");
            col.className = "col-md-1"
            col.innerHTML = testcase.getAttribute("time") + '<span class="text-muted">&nbsp;sec</span>';
            row.appendChild(col);
            col = document.createElement("dd");
            col.className = "col-md-4"
            col.innerHTML = testcase.getAttribute("name")
            row.appendChild(col);
            col = document.createElement("dd");
            col.className = "col-md-6"
            col.innerHTML = testcase.hasChildNodes()? '<span class="text-danger">' + testcase.children[0].getAttribute("message") + '</span>':'<span class="text-success"></span>'
            row.appendChild(col);
        }
        targetElement.appendChild(row);
    }
}
function appendDlChild(dl, dtHtml, ddHtml){
    let dt = document.createElement("dt");
    dt.className = "col-md-1"
    dt.innerHTML = dtHtml;
    dl.appendChild(dt);
    let dd = document.createElement("dd");
    dd.className = "col-md-1"
    dd.innerHTML = ddHtml;
    dl.appendChild(dd);
}