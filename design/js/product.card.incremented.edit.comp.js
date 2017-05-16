/**
 * Created by Julius Alvarado on 5/15/2017.
 */

(function(){
    "use strict";

    var app = angular.module('ap-slider'),
        componentId = 'productCardIncrementedEdit';

    app.component(componentId, {
        templateUrl: 'design/js/product.card.incremented.edit.temp.html',
        bindings: {
            product: '<'
        },
        controller: [IncrementedProductGroupCtrl]
    });

    function IncrementedProductGroupCtrl () {
        var vm = this;
        vm.showImageUrl = true;
        vm.showHeader = true;
        vm.showProductId = true;
        vm.showPrice = true;
        vm.buttonClicked = false;
        vm.buttonText = vm.buttonClicked ? 'Save/Cancel' : 'Edit';
        vm.message = "Incremented Product Cards";

        vm.setEachProduct = function(index, product){
            // console.log("index = "+index);
            // console.log("product.name = "+product.name);
        };

        vm.editCard = function(){
            vm.buttonClicked = !vm.buttonClicked;
            vm.buttonText = vm.buttonClicked ? 'Save/Cancel' : 'Edit';
            vm.showImageUrl = !vm.showImageUrl;
            vm.showHeader = !vm.showHeader;
            vm.showProductId = !vm.showProductId;
            vm.showPrice = !vm.showPrice;

            vm.product.image = vm.productImgUrl ? vm.productImgUrl : vm.product.image;
            vm.product.name = vm.productTitle ? vm.productTitle : vm.product.name;
            vm.product.productId = vm.productId ? vm.productId : vm.product.productId;
            vm.product.price = vm.productPrice ? vm.productPrice : vm.product.price;
            //TODO: Add validation to this setter.
            vm.product.$id = vm.productId ? vm.productId : vm.product.productId;
            // vm.product.name = vm.productTitle;

            if( vm.productImgUrl || vm.productTitle || vm.productId || vm.productPrice ) {
                jProductGroup1Data.Row1Group1.$save(vm.product);
            }
        };

        vm.$onInit = function(){

        };
    }
})();