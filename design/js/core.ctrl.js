/**
 * Created by user on 4/21/2017.
 */

(function () {
    var app = angular.module('ap-slider', ['firebase']);

    app.controller('apCoreCtrl', ['$scope', '$firebaseObject', '$firebaseArray',
        function ($scope, $firebaseObject, $firebaseArray) {
            var ref_productData = firebase.database().ref().child('productData');
            var ref_messages = firebase.database().ref().child('messages');
            var syncObject = $firebaseObject(ref_productData);

            $scope.message = "Ello World ^_^/";
            $scope.messages = $firebaseArray(ref_messages);
            $scope.addMessage = function () {
                $scope.messages.$add({
                    text: $scope.newMessageText,
                    zdate: Date.now()
                });
                $scope.newMessageText = '';
            };

            //-- 3 way data binding:
            syncObject.$bindTo($scope, 'data');
        }
    ]);
})();


//