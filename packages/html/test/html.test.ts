import assert from 'node:assert/strict'
import { test } from 'node:test'
import { createKernel } from '../../kernel/src/index'
import { bindHtmlTemplate, HtmlAdapter } from '../src/index'

test('HtmlAdapter loads an HTML string and discovers fields', async () => {
  const kernel = createKernel({ adapter: new HtmlAdapter() })
  await kernel.dispatch({
    type: 'load',
    source: {
      kind: 'html',
      html: '<main><h1>{{title}}</h1><p>{{amount:number}}</p></main>',
    },
  })

  assert.deepEqual(
    kernel.getFormSchema().fields.map(({ name, type }) => ({ name, type })),
    [
      { name: 'title', type: 'text' },
      { name: 'amount', type: 'number' },
    ],
  )
  assert.equal(kernel.getDefinition()?.source.kind, 'html')
})

test('HtmlAdapter loads UTF-8 bytes and exports bound HTML', async () => {
  const kernel = createKernel({ adapter: new HtmlAdapter() })
  await kernel.dispatch({
    type: 'load',
    source: {
      kind: 'html',
      buffer: new TextEncoder().encode('<p>甲方：{{partyA}}</p>'),
    },
  })
  await kernel.dispatch({ type: 'setValue', path: 'partyA', value: '<甲公司>' })
  const result = await kernel.dispatch({ type: 'export' })
  assert.equal(result.type, 'exported')
  if (result.type === 'exported') {
    assert.equal(result.format, 'html')
    assert.equal(new TextDecoder().decode(result.buffer), '<p>甲方：&lt;甲公司&gt;</p>')
  }
})

test('bindHtmlTemplate expands repeating table rows', () => {
  const result = bindHtmlTemplate(
    '<table><tr><td>{{items.$index}}</td><td>AT&amp;T {{items.name}}</td></tr></table>',
    { items: [{ name: '<A>' }, { name: 'B' }] },
  )
  assert.equal(
    result,
    '<table><tr><td>1</td><td>AT&amp;T &lt;A&gt;</td></tr><tr><td>2</td><td>AT&amp;T B</td></tr></table>',
  )
})
