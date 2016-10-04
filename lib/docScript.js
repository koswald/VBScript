
var docScript = { 

    firstClick: {}, // use firstClick to prevent nothing happening on the first click

    scrollBy: 0, // pixels to scrolldown on click when showing detail

    // toggle visibility of details

    toggleDetail: function(e) {
        var id = e.target.id || window.event.srcElement.id;
        if ("" == id) return;
        var detail = document.getElementsByClassName("detail")[id];

        if (undefined === this.firstClick[id]) this.firstClick[id] = true;

        if (detail.style.display == 'none' || this.firstClick[id]) {

            detail.style.display = 'block'; // show detail; use 'block' here if the .css specifies none (hidden)

            window.scrollBy(0, this.scrollBy);

        } else {
            detail.style.display = 'none';  // hide the detail
        }
        this.firstClick[id] = false;
    },

    // show selected output while debugging

    out: function(str) {
        document.getElementsByClassName("debugOutput")[0].innerHTML += ', ' + str;
    }

}