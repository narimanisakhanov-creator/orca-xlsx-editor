// Writes orca-catalog.json from `orca plugin release-entry`.
//
// Release order matters: release-entry reads the committed HEAD, not the working tree, and the
// catalog itself is part of the hashed release content. Therefore:
//
//   1. build and verify all plugin content, including the existing catalog, commit it as A, and tag A;
//   2. with A checked out and no uncommitted changes, run this script;
//   3. commit the generated orca-catalog.json as a separate commit B.
//
// A's existing catalog is included in A's content hash. Do not amend A with the generated entry,
// move the tag, or re-pin the entry to B: B contains that entry, so it would be a self-reference.
// The entry must name A and A's content hash.
//
//   node scripts/write-catalog.mjs
//
// The catalog and the plugin live in the same repository, so generating the catalog before A is
// committed would describe the previous HEAD, not the release being prepared.
import { execFileSync } from 'node:child_process'
import { readFileSync, writeFileSync } from 'node:fs'
import path from 'node:path'

const ROOT = path.resolve(import.meta.dirname, '..')
const ORCA = process.env.ORCA ?? 'orca'
const manifest = JSON.parse(readFileSync(path.join(ROOT, 'orca-plugin.json'), 'utf8'))
const repository = manifest.repository
const owner = new URL(repository).pathname.split('/').filter(Boolean)[0]

// `release-entry` hashes the files of HEAD; uncommitted changes are ignored. Run it only on the
// committed release commit A, never on a tree whose release changes are still uncommitted.
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
