/**
 * Created by Julius Alvarado on 5/10/2017.
 */


(function () {
    var app = angular.module('ap-slider'),
        componentId = 'productCardEdit';
    app.component(componentId, {
        bindings: {
            product: '='
        },
        templateUrl: 'design/js/product.card.temp.html',
        controller: function (jProductGroup1Data) {
            var vm = this;
            vm.showImageUrl = true;
            vm.showHeader = true;
            vm.showProductId = true;
            vm.showPrice = true;

            vm.editHeader = function () {
                vm.showPrice = true;
                vm.showProductId = true;
                vm.showImageUrl = true;
                vm.showHeader = !vm.showHeader;
            };

            vm.editProductId = function () {
                vm.showHeader = true;
                vm.showPrice = true;
                vm.showImageUrl = true;
                vm.showProductId = !vm.showProductId;
            };

            vm.editPrice = function () {
                vm.showHeader = true;
                vm.showProductId = true;
                vm.showImageUrl = true;
                vm.showPrice = !vm.showPrice;
            };

            vm.editImageUrl = function () {
                vm.showHeader = true;
                vm.showProductId = true;
                vm.showPrice = true;
                vm.showImageUrl = !vm.showImageUrl;
            };
        }
    });
})();