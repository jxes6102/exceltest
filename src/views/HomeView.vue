<template>
    <div class="w-[100vw] h-[100vh] flex flex-wrap justify-center items-start">
        <div class="w-full mt-4 flex flex-wrap justify-center items-center gap-y-2">
            <div class="w-full flex flex-wrap justify-center items-center gap-y-2">
                <div class="w-full md:w-1/2 flex flex-wrap justify-center md:justify-end items-center">
                    <span class="mx-2 text-base font-medium text-gray-900">輸入檔案名稱</span>
                    <input disabled v-model="inputData.fileName" type="text" class="w-1/2 md:w-1/3 p-2 text-gray-900 border border-gray-300 rounded-lg bg-gray-50 text-base focus:ring-blue-500 focus:border-blue-500 ">
                </div>
                <div class="w-full md:w-1/2 pl-2 flex flex-wrap justify-center md:justify-start items-center">
                    <span class="mx-2 text-base font-medium text-gray-900 ">折扣</span>
                    <select v-model="inputData.discount" class="w-1/2 md:w-1/3 bg-gray-50 border border-gray-300 text-gray-900 text-base rounded-lg focus:ring-blue-500 focus:border-blue-500 block p-2.5 ">
                        <option v-for="(item,index) in discountOption" :key="index" :value="item.value">{{item.name}}</option>
                    </select>
                </div>
            </div>
            <div class="w-full flex flex-wrap justify-center items-center gap-y-2">
                <div class="w-full md:w-1/2 flex flex-wrap justify-center md:justify-end items-center">
                    <span class="mx-2 text-base font-medium text-gray-900 ">品牌欄位</span>
                    <select v-model="inputData.brandColumn" class="w-1/2 md:w-1/3 bg-gray-50 border border-gray-300 text-gray-900 text-base rounded-lg focus:ring-blue-500 focus:border-blue-500 block p-2.5 ">
                        <option v-for="(item,index) in columnList" :key="index" :value="item.value">{{item.name}}</option>
                    </select>
                </div>
                <div class="w-full md:w-1/2 pl-2 flex flex-wrap justify-center md:justify-start items-center">
                    <span class="mx-2 text-base font-medium text-gray-900 ">價格欄位</span>
                    <select v-model="inputData.priceColumn" class="w-1/2 md:w-1/3 bg-gray-50 border border-gray-300 text-gray-900 text-base rounded-lg focus:ring-blue-500 focus:border-blue-500 block p-2.5">
                        <option v-for="(item,index) in columnList" :key="index" :value="item.value">{{item.name}}</option>
                    </select>
                </div>
                
                
            </div>
            <!-- <div class="w-full flex flex-wrap justify-center items-center">
                <span class="mx-2 text-base font-medium text-gray-900 ">品牌</span>
                <select multiple v-model="inputData.brand" class="w-3/4 bg-gray-50 border border-gray-300 text-gray-900 text-base rounded-lg focus:ring-blue-500 focus:border-blue-500 block w-full p-2.5 ">
                    <option v-for="(item,index) in brandOptions" :key="index" :value="item.value">{{item.name}}</option>
                </select>
            </div> -->
            
            
        </div>
        <div class="w-full flex flex-wrap justify-center items-center">
            <div class="drop-zone w-[240px] h-[120px] md:w-[500px] md:h-[250px] flex-col mine-flex-center cursor-pointer border-[#009578] border-4 rounded-2xl bg-[#e0ffb5]" ref="fileDiv" :class="[borderStyle ? 'border-solid' : 'border-dashed',]"
                @click="choseFile"><span>匯入xlsx檔案</span><input class="drop-zone__input hidden" ref="fileInput" type="file" name="myFile" @change="changeFile" />
            </div>
        </div>
        
        <div class="w-full flex flex-wrap justify-center items-center gap-2">
            <button class="bg-blue-500 hover:bg-blue-700 text-white font-bold py-2 px-4 rounded" @click="download">download</button>
            <button class="bg-blue-500 hover:bg-blue-700 text-white font-bold py-2 px-4 rounded"  @click="clear">clear</button>
        </div>
    </div>
</template>

<script setup>
import { getLanguage, setLanguage } from "@/lang"
import {useCounterStore} from '../stores/counter'
import { ref,computed } from "vue";
import { useRouter,useRoute } from "vue-router";
import { useI18n } from 'vue-i18n'
import * as XLSX from 'xlsx/xlsx.mjs'
// import fs from 'vite-plugin-fs/browser';
import { writeFile } from "xlsx";

// const { t,locale } = useI18n()

let allData = null
const fileInput = ref(null)
const fileDiv = ref(null)
const borderStyle = ref(false)
const discountOption = [
    {name:'1折',value:0.1},
    {name:'15折',value:0.15},
    {name:'2折',value:0.2},
    {name:'25折',value:0.25},
    {name:'3折',value:0.3},
    {name:'35折',value:0.35},
    {name:'4折',value:0.4},
    {name:'45折',value:0.45},
    {name:'5折',value:0.5},
    {name:'55折',value:0.55},
    {name:'6折',value:0.6},
    {name:'65折',value:0.65},
    {name:'7折',value:0.7},
    {name:'75折',value:0.75},
    {name:'8折',value:0.8},
    {name:'85折',value:0.85},
    {name:'9折',value:0.9},
    {name:'95折',value:0.95},
]
const columnList = [
    {name:'A',value:'A'},
    {name:'B',value:'B'},
    {name:'C',value:'C'},
    {name:'D',value:'D'},
    {name:'E',value:'E'},
    {name:'F',value:'F'},
    {name:'G',value:'G'},
    {name:'H',value:'H'},
    {name:'I',value:'I'},
    {name:'J',value:'J'},
    {name:'K',value:'K'},
    {name:'L',value:'L'},
    {name:'M',value:'M'},
    {name:'N',value:'N'},
    {name:'O',value:'O'},
    {name:'P',value:'P'},
    {name:'Q',value:'Q'},
    {name:'R',value:'R'},
    {name:'S',value:'S'},
    {name:'T',value:'T'},
    {name:'U',value:'U'},
    {name:'V',value:'V'},
    {name:'W',value:'W'},
    {name:'X',value:'X'},
    {name:'Y',value:'Y'},
    {name:'Z',value:'Z'},
]
const inputData = ref({
    fileName:'',
    brandColumn:'B',
    priceColumn:'D',
    discount:0.65,
    brand:['NFB','ELCB','MS','MC']
})
// const brandOptions = [
//   {
//     value: 'NFB',
//     name: 'NFB',
//   },
//   {
//     value: 'ELCB',
//     name: 'ELCB',
//   },
//   {
//     value: 'MS',
//     name: 'MS',
//   },
//   {
//     value: 'MC',
//     name: 'MC',
//   },
// ]

const choseFile = () => {
    fileInput.value.click()
}

const changeFile = () => {
    if(fileInput.value.files.length === 1) {
        let fileType = fileInput.value.files[0].name.split('.')[1]
        if(fileType == 'xlsx'){
            dealFile(fileInput.value.files[0])
        }
    }
}

const dealFile = (file) => {
    inputData.value.fileName = file.name
    const reader = new FileReader()
    reader.readAsArrayBuffer(file)
    reader.onload = () => {
        const data = new Uint8Array(reader.result)
        const wb = XLSX.read(data, {type:'array'})
        allData = wb
        // console.log('allData',allData)
        
    }
}

const clear = () => {
    //console.log('fileInput.value.files',fileInput.value.files,typeof fileInput.value.files)
    fileInput.value.value = ''
    inputData.value.fileName = ''
    allData = null
}

const getSheetName = () => {
    //只取第一張表
    return allData['SheetNames'][0]
}

const getRange = (data) => {
    let temp = data.split(':')
    return {start:temp[0],end:temp[1]}
}

const checkNumber = (value) => {
    if (!isNaN(parseFloat(value)) && isFinite(value)) {
        return true
    } else {
        return false
    }
}

//NFB ELCB MS MC
const checkBrand = (name) => {
    if(typeof name != "string"){
        return false
    }
    let brand = ['NFB','ELCB','MS','MC']
    let value = name.trim()
    return brand.includes(value) 
}


const download = () => {
    if(!allData){
        return
    }

    let sheetNames = getSheetName()
    let sheetData = allData.Sheets[sheetNames]
    let range = getRange(allData.Sheets[sheetNames]['!ref'])

    let allCell =  Object.keys(sheetData) 

    for(let i = 0;i<allCell.length;i++){
        if(allCell[i]!=="!ref" || allCell[i]!=="!margins"){
            let brandCell = sheetData[allCell[i]]
            if(allCell[i].includes(inputData.value.brandColumn) && checkBrand(brandCell.v)){
                let numberStr = allCell[i].replace(/\D/g, "")
                // const lettersOnly = str.replace(/[^a-zA-Z]/g, "");
                let targetCellName = inputData.value.priceColumn+numberStr

                if(allCell.includes(targetCellName)){
                    let targetCell = sheetData[targetCellName]

                    if(checkNumber(targetCell.v)){
                        targetCell.v = Math.ceil(parseInt(targetCell.v)*inputData.value.discount);
                    }

                    if(checkNumber(targetCell.w)){
                        targetCell.w = Math.ceil(parseInt(targetCell.w)*inputData.value.discount);
                    }

                    if(checkNumber(targetCell.v)){
                        targetCell.w = Math.ceil(parseInt(targetCell.w)*inputData.value.discount);
                    }
                    
                }
                
            }
        }
    }
    
    let name =  inputData.value.fileName.split('.')[0]
    /* output format determined by filename */
    writeFile(allData, (name+"修改.xlsx"));

}
</script>

<style scoped>
:deep(.el-select__wrapper){
    width: 150px;
}

@media screen and (min-width: 768px) {
    :deep(.el-select__wrapper){
        width: 400px;
    }
}


</style>