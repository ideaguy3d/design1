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
            vm.buttonClicked = false;
            vm.buttonText = vm.buttonClicked ? 'Save' : 'Edit';

            vm.editCard = function(){
                vm.buttonClicked = !vm.buttonClicked;
                vm.buttonText = vm.buttonClicked ? 'Save' : 'Edit';
                vm.showImageUrl = !vm.showImageUrl;
                vm.showHeader = !vm.showHeader;
                vm.showProductId = !vm.showProductId;
                vm.showPrice = !vm.showPrice;

                vm.product.image = vm.productImgUrl ? vm.productImgUrl : vm.product.image;
                vm.product.name = vm.productTitle ? vm.productTitle : vm.product.name;
                vm.product.productId = vm.productId ? vm.productId : vm.product.productId;
                vm.product.price = vm.productPrice ? vm.productPrice : vm.product.price;
                // vm.product.name = vm.productTitle;

                if( vm.productImgUrl || vm.productTitle || vm.productId || vm.productPrice ) {
                    jProductGroup1Data.Row1.$save(vm.product);
                }
            };

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
            vm.updateImageUrl = function(){
                console.log("productId = "+vm.product.image);
            };

        }
    });
})();