/**
 * Created by Julius Alvarado on 5/15/2017.
 */
(function(){
    var app = angular.module('ap-slider'),
        componentId = 'productItems';

    app.component(componentId, {
        templateUrl: 'design/js/product.items.temp.html',
        bindings: {
            items: '<'
        },
        controller: [ItemsCtrl]
    });

    function ItemsCtrl () {

    }
})();