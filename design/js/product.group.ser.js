/**
 * Created by Julius Alvarado on 5/5/2017.
 */

(function(){

    var app = angular.module('ap-slider');
    // products group 1 data service
    app.factory("jProductGroup1Data", ['$firebaseObject', '$firebaseArray',
        function ($firebaseObject, $firebaseArray) {
            var ref_messages = firebase.database().ref().child('messages');
            var ref_row1 = firebase.database().ref().child('Row1');
            var ref_row1_group1 = firebase.database().ref().child('Row1').child('Group1');
            var ref_row1_group2 = firebase.database().ref().child('Row1').child('Group2');
            var ref_jcategories = firebase.database().ref().child('jcategories').orderByChild("name");

            return {
                Row1: $firebaseArray(ref_row1),
                Row1Group1: $firebaseArray(ref_row1_group1),
                Row1Group2: $firebaseArray(ref_row1_group2),
                jcategories: $firebaseArray(ref_jcategories),
                ProductsMessagesArray: function () {
                    return $firebaseArray(ref_messages);
                }
            }
        }
    ]);
})();
