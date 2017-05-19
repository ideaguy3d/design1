/**
 * Created by Julius Alvarado on 5/17/2017.
 */
(function () {
    var app = angular.module('ap-slider'),
        componentId = 'productCard';

    app.component(componentId, {
        bindings: {
            product: '=',
            activeArea: '<'
        },
        templateUrl: 'design/js/comp.temp.product.card.html',
        controller: ['jProductGroup1Data', ProductCardCtrl]
    });

    function ProductCardCtrl(jProductGroup1Data) {
        var vm = this;
    } // END OF:  function ProductCardEditCtrl(); ^_^/

})();