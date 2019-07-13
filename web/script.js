function start() {
    document.getElementById('status_text').innerHTML = "Connecting..."
    eel.meas_test();
}

function progress_two() {
    console.log("hello from advance progress");
    var stepNum = document.getElementById('nr').value;
    var p = stepNum * 16.6;
    if (stepNum === 0) {
        let p = 0
    };
    document.getElementsByClassName('percent')[0].style.width = `${p}%`;
    steps.forEach((e) => {
        if (e.id === stepNum) {
            e.classList.add('selected');
            e.classList.add('completed');
        }
        if (e.id < stepNum) {
            e.classList.add('completed');
        }
        if (e.id > stepNum) {
            e.classList.remove('selected', 'completed');
        }
    });
}

eel.expose(changeProgress);
function changeProgress(nr) {
    document.getElementById('nr').value = nr;
    if (nr === 1) {
        document.getElementById('status_text').innerHTML = "Connecting..."
    } else if (nr === 2) {
        document.getElementById('status_text').innerHTML = "Voltage measurements..."
    } else if (nr === 3) {
        document.getElementById('status_text').innerHTML = "Switch position"
    } else if (nr === 4) {
        document.getElementById('status_text').innerHTML = "Z, R, C measurements..."
    } else if (nr === 5) {
        document.getElementById('status_text').innerHTML = "Creating export..."
    } else if (nr === 6) {
        document.getElementById('status_text').innerHTML = "Done!"
        document.getElementById("download_excel").style.display = "block";
    }
    progress_two();
}

eel.expose(alertSwitch);
function alertSwitch() {
    alert("Please move switch position to measure impedance");
}