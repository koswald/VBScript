// used by the documentation .html file

var docScript = { 

    // use firstClick to prevent unresponsive first click

    firstClick: {}, 

    // pixels to scrolldown on click when showing detail

    scrollBy: 0, 

    // toggle visibility of details

    toggleDetail: function(e) {

        var id = e.target.id || window.event.srcElement.id;
        if ("" == id) return;
        var detail = document.getElementsByClassName("detail")[id];

        if (undefined === this.firstClick[id]) this.firstClick[id] = true;

        if (detail.style.display == 'none' || this.firstClick[id]) {

            // show detail; use 'block' here 
            // if the .css specifies none (hidden)

            detail.style.display = 'block';

            window.scrollBy(0, this.scrollBy);

        } else {

            // hide the detail

            detail.style.display = 'none';
        }
        this.firstClick[id] = false;
    },

    // show selected output while debugging

    out: function(str) {
        document.getElementsByClassName("debugOutput")[0].innerHTML += ', ' + str;
    }

}