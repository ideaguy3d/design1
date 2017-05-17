/**
 * Created by Julius Alvarado on 5/15/2017.
 */

(function () {
    var app = angular.module('ap-slider'),
        componentId = 'productGroupIncrementedEdit';

    // component definition
    app.component(componentId, {
        templateUrl: 'design/js/product.group.incremented.temp.html',
        bindings: {
            groupItems: '<',
            activeArea: '<'
        },
        controller: ['jProductGroup1Data', ProductGroupIncrementedCtrl]
    });

    function ProductGroupIncrementedCtrl(jProductGroup1Data) {
        var vm = this;
        vm.divGroupItems = [];
        vm.apcRow1Group2 = jProductGroup1Data.Row1Group2;
        vm.apcRow1Group3 = jProductGroup1Data.Row1Group3;
        vm.apcRow1Group4 = jProductGroup1Data.Row1Group4;

        vm.setGroupItem = function (index, groupItem) {
            console.log(groupItem);
            for (var k in groupItem) {
                // console.log("index = "+index+", key = "+k);
            }
        };
    }
})();