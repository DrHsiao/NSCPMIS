function changeCityValue(CityValue) {
    //顯示特定城市數據
    var townContainers = document.getElementsByClassName('town-container');
    for (var i = 0; i < townContainers.length; i++) {
        if (townContainers[i].id === CityValue) {
            townContainers[i].style.display = 'block';
        } else {
            townContainers[i].style.display = 'none';
        }
    }
}

function handleCityCheckbox(checkbox, CityValue) {
    var city = checkbox.getAttribute("data-city");
    var townContainer = document.getElementById(city);
    var townCheckboxes = townContainer.querySelectorAll('input[type="checkbox"]');

    for (var i = 0; i < townCheckboxes.length; i++) {
        townCheckboxes[i].checked = checkbox.checked;
    }

    var townContainers = document.getElementsByClassName('town-container');
    for (var i = 0; i < townContainers.length; i++) {
        if (townContainers[i].id === CityValue) {
            townContainers[i].style.display = 'block';
        } else {
            townContainers[i].style.display = 'none';
        }
    }
}

function toggleCheckboxes(city) {
    var townContainer = document.getElementById(city);
    var townCheckboxes = townContainer.querySelectorAll('input[type="checkbox"]');

    for (var i = 0; i < townCheckboxes.length; i++) {
        townCheckboxes[i].checked = !townCheckboxes[i].checked;
    }

    var selectAllButton = document.getElementById("selectAll");
    var allTownCheckboxes = document.querySelectorAll('.town-container input[type="checkbox"]');
    var allChecked = true;

    for (var i = 0; i < allTownCheckboxes.length; i++) {
        if (!allTownCheckboxes[i].checked) {
            allChecked = false;
            break;
        }
    }

    selectAllButton.textContent = allChecked ? "取消全選" : "全選";
}