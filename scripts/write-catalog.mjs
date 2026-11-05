// Writes orca-catalog.json from the manifest and `orca plugin release-entry`, so the catalog can
// never drift from the plugin it lists. Run after `npm run build` and before a release commit.
//
//   node scripts/write-catalog.mjs
//
// The catalog and the plugin live in the same repository: the catalog is not part of the plugin's
// content hash, because Orca hashes what the manifest ships, not every file in the folder.
import { execFileSync } from 'node:child_process'
import { readFileSync, writeFileSync } from 'node:fs'
import path from 'node:path'

const ROOT = path.resolve(import.meta.dirname, '..')
const ORCA = process.env.ORCA ?? 'orca'
const manifest = JSON.parse(readFileSync(path.join(ROOT, 'orca-plugin.json'), 'utf8'))
const repository = manifest.repository
const owner = new URL(repository).pathname.split('/').filter(Boolean)[0]

const raw = execFileSync(ORCA, ['plugin', 'release-entry', '.', '--json'], {
  cwd: ROOT,
  encoding: 'utf8'
})
const entry = JSON.parse(raw.slice(raw.indexOf('{'))).result.entry

const catalog = {
  $schema: 'https://raw.githubusercontent.com/stablyai/orca/main/docs/schemas/plugins/orca-catalog.v2.schema.json',
  catalogVersion: 2,
  name: manifest.name,
  owner,
  plugins: [
    {
      id: `${owner}.${manifest.id}`,
      source: { kind: 'git', url: repository, ref: 'main' },
      description: manifest.description,
      categories: ['editors'],
      releases: [
        {
          version: entry.version,
          commit: entry.commit,
          contentHash: entry.contentHash,
          execution: entry.execution,
          minimumOrca: entry.minimumOrca,
          name: entry.name,
          license: manifest.license,
          contributes: entry.contributes,
          fileTypes: entry.fileTypes,
          editorFileTypes: entry.editorFileTypes,
          images: entry.images
        }
      ]
    }
  ]
}

writeFileSync(path.join(ROOT, 'orca-catalog.json'), `${JSON.stringify(catalog, null, 2)}\n`)
console.log(`orca-catalog.json: ${catalog.plugins[0].id} ${entry.version} ${entry.commit.slice(0, 8)}`)
