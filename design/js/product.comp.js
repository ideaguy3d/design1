/**
 * Created by Julius Alvarado on 5/5/2017.
 */

(function(){
    var app = angular.module('ap-slider'),
        componentId = 'sliderProduct';

    app.component(componentId, {
        templateUrl: 'design/views/templates/slider.product.tem.html',
        controller: function(jProductGroup1Data) {
            var vm = this;
            vm.jcategories = jProductGroup1Data.jcategories;
            vm.categoriesTitle = "AussieProducts.com categories";

            vm.createNewCategory = function () {
                vm.jcategories.$add({name: vm.newCategoryName});
                vm.newCategoryName = '';
            }
        }
    });
})();
