import axios from 'axios'
import request from '@/utils/request'
import { getToken } from '@/utils/auth'

const baseUrl = import.meta.env.VITE_APP_BASE_API
const formClient = axios.create()
const jsonClient = axios.create()

delete formClient.defaults.headers.common['Content-Type']
delete formClient.defaults.headers.post['Content-Type']
delete formClient.defaults.headers.put['Content-Type']
delete formClient.defaults.headers.patch['Content-Type']

function postExcelForm(url, formData) {
  const headers = {}
  const token = getToken()
  if (token) {
    headers.Authorization = `Bearer ${token}`
  }
  return formClient.post(`${baseUrl}${url}`, formData, {
    headers,
    transformRequest: [(data) => data]
  }).then((response) => {
    const result = response.data || {}
    if (result.code !== 200) {
      return Promise.reject(new Error(result.msg || 'Request failed'))
    }
    return result
  })
}

function requestExcelJson(method, url, data, params) {
  const headers = {}
  const token = getToken()
  if (token) {
    headers.Authorization = `Bearer ${token}`
  }
  return jsonClient({
    method,
    url: `${baseUrl}${url}`,
    data,
    params,
    headers
  }).then((response) => response.data || {})
}

export function getExcelEditorInfo() {
  return request({
    url: '/tool/excel-editor/info',
    method: 'get'
  })
}

export function getExcelEditorContent(fileName) {
  return request({
    url: '/tool/excel-editor/content',
    method: 'get',
    params: fileName ? { fileName } : undefined,
    responseType: 'arraybuffer'
  })
}

export function getExcelWorkbookView(fileName) {
  return request({
    url: '/tool/excel-editor/view',
    method: 'get',
    params: fileName ? { fileName } : undefined
  })
}

export function uploadExcelEditorFile(formData) {
  return postExcelForm('/tool/excel-editor/upload', formData)
}

export function saveExcelEditorFile(formData) {
  return postExcelForm('/tool/excel-editor/save', formData)
}

export function applyExcelWorkbookChanges(data) {
  return requestExcelJson('post', '/tool/excel-editor/apply', data)
}

export function acquireExcelEditorLock(data) {
  return requestExcelJson('post', '/tool/excel-editor/lock', data)
}

export function heartbeatExcelEditorLock(data) {
  return requestExcelJson('post', '/tool/excel-editor/lock/heartbeat', data)
}

export function releaseExcelEditorLock(data) {
  return requestExcelJson('post', '/tool/excel-editor/lock/release', data)
}

export function getExcelEditorLockStatus(fileName) {
  return requestExcelJson('get', '/tool/excel-editor/lock', null, fileName ? { fileName } : undefined)
}
