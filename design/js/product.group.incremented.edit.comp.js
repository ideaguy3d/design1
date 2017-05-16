/**
 * Created by Julius Alvarado on 5/15/2017.
 */

(function(){
    var app = angular.module('ap-slider'),
        componentId = 'productGroupIncrementedEdit';

    app.component(componentId, {
        templateUrl: 'design/js/product.group.incremented.temp.html',
        bindings: {
           groupItems: '<'
        },
        controller: [ProductGroupIncrementedCtrl]
    });

    function ProductGroupIncrementedCtrl () {
        var vm = this;
        vm.setGroupItem = function(index, groupItem){
            for(var k in groupItem) {
                // console.log("index = "+index+", key = "+k);
            }
        };
    }
})();