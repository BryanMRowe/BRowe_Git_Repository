/* Sets Map Window */
var MyMap = L.map("map",{
    center: [35,-100],
    zoom: 6
});

/* Adds Map Layer */
L.tileLayer(
    "https://api.mapbox.com/styles/v1/mapbox/outdoors-v10/tiles/256/{z}/{x}/{y}?"
    +"access_token=pk.eyJ1IjoiZHNhYjg0IiwiYSI6ImNqaWNleGhycTAwYWEzcW1hZzNiYW9wYzMifQ.kelEORaTu-K8YysrN64Csw"
).addTo(myMap);