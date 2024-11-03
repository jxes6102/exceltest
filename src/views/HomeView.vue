<template>
    <div class="w-[100vw] h-[100vh] flex flex-wrap justify-center items-center">
        <div class="w-full flex flex-wrap justify-center items-center gap-2">
            <div class="w-full flex flex-wrap justify-center items-center">
                <span class="mx-2 text-sm font-medium text-gray-900 dark:text-white">輸入檔案名稱</span>
                <input disabled v-model="fileName" type="text" class="w-[200px] p-2 text-gray-900 border border-gray-300 rounded-lg bg-gray-50 text-base focus:ring-blue-500 focus:border-blue-500 dark:bg-gray-700 dark:border-gray-600 dark:placeholder-gray-400 dark:text-white dark:focus:ring-blue-500 dark:focus:border-blue-500">
            </div>
        </div>
        <div class="w-full flex flex-wrap justify-center items-center">
            <div class="drop-zone w-[200px] h-[200px] md:w-[300px] md:h-[300px] flex-col mine-flex-center cursor-pointer border-[#009578] border-4 rounded-2xl bg-[#e0ffb5]" ref="fileDiv" :class="[borderStyle ? 'border-solid' : 'border-dashed',]"
                @click="choseFile"><span>匯入xlsx檔案</span><input class="drop-zone__input hidden" ref="fileInput" type="file" name="myFile" @change="changeFile" />
            </div>
            <!-- <div class="drop-zone w-[200px] h-[200px] md:w-[300px] md:h-[300px] flex-col mine-flex-center cursor-pointer border-[#009578] border-4 rounded-2xl bg-[#e0ffb5]" ref="fileDiv" :class="[borderStyle ? 'border-solid' : 'border-dashed',]"
                @click="choseFile" @drop="dropFile" @dragover="dragOver" @dragleave="setBorder" @dragend="setBorder"><span>匯入檔案</span><span>xlsx檔(無合併儲存格)</span><input class="drop-zone__input hidden" ref="fileInput" type="file" name="myFile" @change="changeFile" />
            </div> -->
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
import HelloWorld from '@/components/HelloWorld.vue'
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
const fileName = ref('')

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

// const dropFile = (e) => {
//     e.preventDefault()
//     borderStyle.value = false
//     if (e.dataTransfer.files.length === 1) {
//         fileInput.value.files = e.dataTransfer.files
//         dealFile(e.dataTransfer.files[0])
//     }
// }
// const dragOver = (e) => {
//     // 未知 未用時會另開頁面
//     e.preventDefault()
//     borderStyle.value = true
// }
// const setBorder = () => {
//     borderStyle.value = false
// }

const dealFile = (file) => {
    fileName.value = file.name
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
    fileName.value = ''
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
            if(allCell[i].includes("B") && checkBrand(brandCell.v)){
                let numberStr = allCell[i].replace(/\D/g, "")
                // const lettersOnly = str.replace(/[^a-zA-Z]/g, "");
                let targetCellName = 'D'+numberStr

                if(allCell.includes(targetCellName)){
                    let targetCell = sheetData[targetCellName]

                    if(checkNumber(targetCell.v)){
                        targetCell.v = Math.ceil(parseInt(targetCell.v)*0.65);
                    }

                    if(checkNumber(targetCell.w)){
                        targetCell.w = Math.ceil(parseInt(targetCell.w)*0.65);
                    }

                    if(checkNumber(targetCell.v)){
                        targetCell.w = Math.ceil(parseInt(targetCell.w)*0.65);
                    }
                    
                }
                
            }
        }
    }
   
    /* output format determined by filename */
    writeFile(allData, "輸出檔案.xlsx");

}
</script>